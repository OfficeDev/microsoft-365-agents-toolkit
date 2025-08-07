using {{YOUR_NAMESPACE}}.Configuration;
using Json.More;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.Dialogs;
using Microsoft.Agents.Builder.State;
using Microsoft.Agents.Core.Models;
using Microsoft.Agents.Storage;
using Microsoft.Extensions.Options;

namespace {{YOUR_NAMESPACE}}.SSO;

public class DialogConstants
{
    internal static readonly string COMMAND_ROUTE_DIALOG = "CommandRouteDialog";
    internal static readonly string TEAMS_SSO_PROMPT = "TeamsFxSsoPrompt";
}

public class StoreItem
{
    public string eTag;
}

public class SsoDialog : ComponentDialog
{
    ILogger<SsoDialog> _logger;
    BotAuthenticationOptions _botAuthOptions;
    Dictionary<string, string> _commandMapping = new Dictionary<string, string>();
    IStorage _dedupStorage = new MemoryStorage();
    List<string> _dedupStorageKeys = new List<string>() {};

    public SsoDialog(IOptions<BotAuthenticationOptions> botAuthenticationOptions, ILogger<SsoDialog> logger)
    {
        _logger = logger;
        InitialDialogId = DialogConstants.COMMAND_ROUTE_DIALOG;

        try
        {
            _logger.LogTrace("Validate bot authentication configuration");
            _botAuthOptions = botAuthenticationOptions.Value;
        }
        catch (OptionsValidationException e)
        {
            throw new Exception($"Bot authentication config is missing or not correct with error: {e.Message}");
        }

        var settings = new TeamsBotSsoPromptSettings(_botAuthOptions, new string[] { "User.Read" });
        AddDialog(new TeamsBotSsoPrompt(DialogConstants.TEAMS_SSO_PROMPT, settings));

        WaterfallDialog commandRouteDialog = new WaterfallDialog(
            DialogConstants.COMMAND_ROUTE_DIALOG,
            new WaterfallStep[]
            {
                CommandRouteStepAsync
            });
        AddDialog(commandRouteDialog);

        _logger.LogInformation("Construct Main Dialog");
    }

    public async Task RunAsync(ITurnContext context, IStatePropertyAccessor<DialogState> accessor)
    {
        DialogSet dialogSet = new DialogSet(accessor);
        dialogSet.Add(this);

        DialogContext dialogContext = await dialogSet.CreateContextAsync(context);
        DialogTurnResult results = await dialogContext.ContinueDialogAsync();
        if (results != null && results.Status == DialogTurnStatus.Empty)
        {
            await dialogContext.BeginDialogAsync(Id);
        }
    }

    public void addCommand(
        string commandId,
        string commandText,
        Func<ITurnContext, string, BotAuthenticationOptions, Task> operation)
    {
        if (_commandMapping.ContainsValue(commandId))
        {
            return;
        }
        _commandMapping.Add(commandText, commandId);

        WaterfallDialog dialog = new WaterfallDialog(
            commandId,
            new WaterfallStep[]
            {
            PromptStepAsync,
            DedupStepAsync,
            async (WaterfallStepContext stepContext, CancellationToken cancellationToken) => {
                TeamsBotSsoPromptTokenResponse tokenResponce = (TeamsBotSsoPromptTokenResponse)stepContext.Result;
                var turnContext = stepContext.Context;
                try
                {
                    if (tokenResponce != null)
                    {
                        await operation(turnContext, tokenResponce.Token, _botAuthOptions);
                    } else
                    {
                        await turnContext.SendActivityAsync("Failed to retrieve user token from conversation context.");
                    }
                    return await stepContext.EndDialogAsync();
                } catch (Exception error)
                {
                    await turnContext.SendActivityAsync("Failed to retrieve user token from conversation context.");
                    await turnContext.SendActivityAsync(error.Message);
                    return await stepContext.EndDialogAsync();
                }
            }
        });
        AddDialog(dialog);
    }

    protected override async Task OnEndDialogAsync(ITurnContext context, DialogInstance instance, DialogReason reason, CancellationToken cancellationToken = default(CancellationToken))
    {
        var conversationId = context.Activity.Conversation.Id;
        var currentDedupKeys = _dedupStorageKeys.Where((key) => key.IndexOf(conversationId) > 0).ToArray();
        await _dedupStorage.DeleteAsync(currentDedupKeys);
        _dedupStorageKeys = _dedupStorageKeys.Where((key) => key.IndexOf(conversationId) < 0).ToList<string>();
    }

    private async Task<DialogTurnResult> PromptStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
    {
        _logger.LogInformation("Step: Prompt to get SSO token");
        try
        {
            return await stepContext.BeginDialogAsync(DialogConstants.TEAMS_SSO_PROMPT, null, cancellationToken);
        } catch (Exception error)
        {
            await stepContext.Context.SendActivityAsync("Failed to run SSO prompt");
            await stepContext.Context.SendActivityAsync(error.Message);
            return await stepContext.EndDialogAsync();
        }
    }

    private async Task<DialogTurnResult> CommandRouteStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
    {
        _logger.LogInformation("Step: Route to pre-added commands");
        var turnContext = stepContext.Context;

        string text = turnContext.Activity.RemoveRecipientMention();
        if (text != null)
        {
            text = text.ToLower().Trim();
        }

        string commandId = MatchCommands(text);
        if (commandId != null)
        {
            return await stepContext.BeginDialogAsync(commandId);
        }

        await stepContext.Context.SendActivityAsync(String.Format("Cannot find command: {0}", text));
        return await stepContext.EndDialogAsync();
    }

    private async Task<DialogTurnResult> DedupStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
    {
        try
        {
            var tokenResponse = stepContext.Result;
            if (tokenResponse != null && (await ShouldDedup(stepContext.Context)))
            {
                return EndOfTurn;
            }
            return await stepContext.NextAsync(tokenResponse);
        }
        catch (Exception error)
        {
            await stepContext.Context.SendActivityAsync("Failed to run dedup step");
            await stepContext.Context.SendActivityAsync(error.Message);
            return await stepContext.EndDialogAsync();
        }
    }

    private async Task<bool> ShouldDedup(ITurnContext context)
    {
        var tokenResponseObject = context.Activity.Value.ToJsonDocument();
        var storeItem = new StoreItem()
        {
            eTag = tokenResponseObject.RootElement.GetProperty("id").ToString(),
        };
        var key = GetStorageKey(context);
        var storeItems = new Dictionary<string, object>()
        {
            {key, storeItem}
        };

        var res = await _dedupStorage.ReadAsync(new string[] { key });
        if (res.Count != 0)
        {
            return true;
        }

        await _dedupStorage.WriteAsync(storeItems);
        _dedupStorageKeys.Add(key);
        return false;
    }

    private string GetStorageKey(ITurnContext context)
    {
        if (context == null || context.Activity == null || context.Activity.Conversation == null)
        {
            throw new Exception("Invalid context, can not get storage key!");
        }

        var activity = context.Activity;
        var channelId = activity.ChannelId;
        var conversationId = activity.Conversation.Id;
        if (activity.Type != ActivityTypes.Invoke || activity.Name != SignInConstants.TokenExchangeOperationName)
        {
            throw new Exception("TokenExchangeState can only be used with Invokes of signin/tokenExchange.");
        }

        var value = activity.Value.ToJsonDocument();
        if (value == null || value.RootElement.GetProperty("id").ToString() == null)
        {
            throw new Exception("Invalid signin/tokenExchange. Missing activity.value.id.");
        }

        return $"{channelId}/{conversationId}/{value.RootElement.GetProperty("id").ToString()}";
    }

    private string MatchCommands(string text)
    {
        if (_commandMapping.ContainsKey(text))
        {
            return _commandMapping[text];
        }

        return null;
    }
}

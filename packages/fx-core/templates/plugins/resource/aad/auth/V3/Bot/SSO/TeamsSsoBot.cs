using Microsoft.Agents.Extensions.Teams.Compat;
using Microsoft.Agents.Builder.Dialogs;
using Microsoft.Agents.Builder.State;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Core.Models;

namespace {{YOUR_NAMESPACE}}.SSO;

public class TeamsSsoBot<T> : TeamsActivityHandler where T : Dialog
{
    private readonly ILogger<TeamsSsoBot<T>> _logger;
    private readonly AgentState _conversationState;
    private readonly Dialog _dialog;
    private readonly IStatePropertyAccessor<DialogState> _dialogState;
    
    public TeamsSsoBot(ConversationState conversationState, T dialog, ILogger<TeamsSsoBot<T>> logger)
    {
        Console.WriteLine("sso bot init");
        _conversationState = conversationState;
        _dialog = dialog;
        _logger = logger;
        _dialogState = _conversationState.CreateProperty<DialogState>("DialogState");

        ((SsoDialog)_dialog).addCommand("showUserInfo", "show", SsoOperations.ShowUserInfo);
    }

    protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
    {
        Console.WriteLine("Receive message activity");
        _logger.LogInformation("Receive message activity");
        await ((SsoDialog)_dialog).RunAsync(turnContext, _dialogState);
    }

    protected override async Task OnTeamsSigninVerifyStateAsync(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
    {
        _logger.LogInformation("Receive invoke activity of teams sign in verify state");
        await _dialog.RunAsync(turnContext, _conversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken);
    }

    protected override async Task OnSignInInvokeAsync(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
    {
        _logger.LogInformation("Receive invoke activity of sign in");
        await _dialog.RunAsync(turnContext, _conversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken);
    }

    public override async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default(CancellationToken))
    {
        await base.OnTurnAsync(turnContext, cancellationToken);
        await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken).ConfigureAwait(false);
    }
}
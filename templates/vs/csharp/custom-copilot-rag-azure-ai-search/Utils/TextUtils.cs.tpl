using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Api.Entities;

namespace {{SafeProjectName}}.Utils
{
    public static class TextUtils
    {
        public static string StripMentionsText(MessageActivity activity)
        {
            if (string.IsNullOrEmpty(activity?.Text))
                return activity?.Text;

            var text = activity.Text;

            var mentions = activity.Entities?
                .OfType<MentionEntity>()
                .ToList() ?? new List<MentionEntity>();

            foreach (var mention in mentions)
            {
                if (!string.IsNullOrEmpty(mention.Text))
                {
                    var textWithoutTags = mention.Text.Replace("<at>", "").Replace("</at>", "");
                    text = text.Replace(mention.Text, string.Empty);
                }
                else if (!string.IsNullOrEmpty(mention.Mentioned.Name))
                {
                    // Handle case where mention.Text is null but we have the mentioned name
                    var mentionPattern = $"<at>{mention.Mentioned.Name}</at>";
                    text = text.Replace(mentionPattern, string.Empty);
                }
            }

            return text.Trim();
        }
    }
}
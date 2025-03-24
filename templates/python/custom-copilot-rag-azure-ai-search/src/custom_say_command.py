import json
from botbuilder.core import MessageFactory, TurnContext
from botframework.connector import Channels
from teams.ai.citations import AIEntity, ClientCitation, Appearance
from teams.utils import snippet
from teams.utils.citations import format_citations_response, get_used_citations
from typing import Dict, Any

async def say_command(context: TurnContext, state: Any, data: Dict[str, Any], feedback_loop_enabled: bool = False) -> str:
    """
    Executes the say command, sending a message to the user with the AI's response.

    Args:
        context (TurnContext): The turn context.
        state (Any): The turn state.
        data (Dict[str, Any]): The data containing the AI's response.
        feedback_loop_enabled (bool): Whether the feedback loop is enabled.

    Returns:
        str: An empty string.
    """
    # Early return if there's no valid response data to process
    if not data or not data.response or not data.response.content:
        return ""

    # Determine if we're in Teams channel for special formatting
    is_teams_channel = context.activity.channel_id == Channels.ms_teams
    content = ""
    result = None

    try:
        # Attempt to parse the response as JSON
        result = json.loads(data.response.content)
    except json.JSONDecodeError as error:
        # If not valid JSON, fallback to sending raw text with AI metadata
        print(f"Response is not valid JSON, sending the raw text. error: {error}")
        message = MessageFactory.text(data.response.content)
        # Add identify this as AI-generated content
        message.entities = [
            AIEntity(
                citation=None,
                additional_type=["AIGeneratedContent"]
            )
        ]
        # Add Teams-specific channel data if needed
        if is_teams_channel:
            message.channel_data = {"feedbackLoopEnabled": feedback_loop_enabled}
        await context.send_activity(message)
        return ""

    # Initialize citations tracking for the response
    citations = []
    position = 1

    # Process JSON result with citations if available
    if result and result.get("results") and len(result["results"]) > 0:
        for content_item in result["results"]:
            if content_item.get("citationTitle") and len(content_item["citationTitle"]) > 0:
                # Create citation object with metadata for the content source
                client_citation = ClientCitation(
                    position = position,
                    appearance = Appearance(
                        name = content_item.get("citationTitle") or f"Document #{position}",
                        url = content_item.get("citationUrl"),
                        # Create a snippet/preview of the citation content, limited to 400 chars(limitation check 480)
                        abstract = snippet(content_item.get("citationContent"), 400),
                    )
                )
                # Add citation reference number to the answer text
                content += f"{content_item.get('answer')}[{position}]<br>"
                position += 1
                citations.append(client_citation)
            else:
                # Add answer without citation if no citation is available
                content += f"{content_item.get('answer')}<br>"
    else:
        # Use raw content if no structured results are available
        content = data.response.content

    # Replace newlines with HTML breaks for Teams rendering
    if is_teams_channel:
        content = content.replace("\n", "<br>")

    # Format the content with proper citation handling
    content_text = format_citations_response(content) if len(citations) < 1 else content
    # Get only the citations that are actually referenced in the text
    referenced_citations = get_used_citations(content_text, citations) if len(citations) > 0 else None

    # Create the message with properly formatted text
    message = MessageFactory.text(content_text)
    # Add AI entity with citation metadata
    message.entities = [
        AIEntity(
            citation=(referenced_citations if referenced_citations else []),
            additional_type=["AIGeneratedContent"],
        ),
    ]

    # Add Teams-specific channel data if needed
    if is_teams_channel:
        message.channel_data = {"feedbackLoopEnabled": feedback_loop_enabled}

    # Send the formatted message to the user
    await context.send_activity(message)
    return ""
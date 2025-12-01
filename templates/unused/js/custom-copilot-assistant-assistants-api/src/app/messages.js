async function resetMessage(context, state) {
  state.deleteConversationState();
  await context.sendActivity("Ok lets start this over.");
}

module.exports = {
  resetMessage,
};

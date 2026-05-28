export class LanguageModelChatMessage {
  /**
   * Utility to create a new user message.
   *
   * @param content The content of the message.
   * @param name The optional name of a user for the message.
   */
  static User(content: string, name?: string): LanguageModelChatMessage {
    return new LanguageModelChatMessage(LanguageModelChatMessageRole.User, content, name);
  }

  /**
   * Utility to create a new assistant message.
   *
   * @param content The content of the message.
   * @param name The optional name of a user for the message.
   */
  static Assistant(content: string, name?: string): LanguageModelChatMessage {
    return new LanguageModelChatMessage(LanguageModelChatMessageRole.Assistant, content, name);
  }

  /**
   * The role of this message.
   */
  role: LanguageModelChatMessageRole;

  /**
   * The content of this message.
   */
  content: string;

  /**
   * The optional name of a user for this message.
   */
  name: string | undefined;

  /**
   * Create a new user message.
   *
   * @param role The role of the message.
   * @param content The content of the message.
   * @param name The optional name of a user for the message.
   */
  constructor(role: LanguageModelChatMessageRole, content: string, name?: string) {
    this.role = role;
    this.content = content;
    this.name = name;
  }
}

export enum LanguageModelChatMessageRole {
  /**
   * The user role.
   */
  User = 1,

  /**
   * The assistant role.
   */
  Assistant = 2,
}

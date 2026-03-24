import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";

export default class CopilotService {
  private context: WebPartContext;
  private client: MSGraphClientV3 | undefined;
  private conversationId: string | null = null;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  private async getClient(): Promise<MSGraphClientV3> {
    if (!this.client) {
      this.client = await this.context.msGraphClientFactory.getClient("3");
    }
    return this.client;
  }

  // ⭐ Create conversation
  public async createConversation(): Promise<string> {
    const client = await this.getClient();

    const res = await client
      .api("/copilot/conversations")
      .version("beta")
      .post({});

    this.conversationId = res.id;

    if (!this.conversationId) {
      throw new Error("Conversation not created");
    }

    console.log("CONVERSATION CREATED →", this.conversationId);

    return this.conversationId;
  }

  // ⭐ Send message
  public async sendMessage(prompt: string, fileUrl?: string | null): Promise<string> {
    const client = await this.getClient();

    if (!this.conversationId) {
      await this.createConversation();
    }

    const finalPrompt = fileUrl
      ? `${prompt}\n\nFile: ${fileUrl}`
      : prompt;

    console.log("SENDING TO COPILOT →", finalPrompt);

    const sendResult = await client
      .api(`/copilot/conversations/${this.conversationId}/chat`)
      .version("beta")
      .post({
        locationHint: {
          timeZone: "Asia/Kolkata"
        },
        message: {
          text: finalPrompt
        }
      });

    console.log("COPILOT RESPONSE →", sendResult);

    const answer = sendResult?.messages?.[1]?.text;

    return answer || "Copilot did not respond";
  }
}
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";

export default class CopilotService {
  private context: WebPartContext;
  private client: MSGraphClientV3 | undefined;
  private conversationId: string | null = null;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  private async getClient() {
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

  // ⭐ Send message + debug execution
  public async sendMessage(prompt: string,fileUrl?: string | null): Promise<string> {

    const client = await this.getClient();

    if (!this.conversationId) {
      await this.createConversation();
    }
    
    const finalPrompt = fileUrl
      ? `${prompt}\n\nFile: ${fileUrl}`
      : prompt;

    console.log("SENDING TO COPILOT →", {
      conversationId: this.conversationId,
      prompt: finalPrompt
    });

    // ⭐ STEP 2 — Send message and log response
    const sendResult = await client
      .api(`/copilot/conversations/${this.conversationId}/chat`)
      .version("beta")
      .post({
        // role: "user",
        // content: [
        //   {
        //     type: "text",
        //     text: finalPrompt
        //   }
        // ],
        locationHint: {
          timeZone: "America/New_York"
        },
        message: {
          text: finalPrompt
        },
      });

    console.log("COPILOT SEND RESULT →", sendResult);

    const answer = sendResult?.messages?.[1]?.text;

     const question = sendResult.messages[0]?.text;
      const answerr = sendResult.messages[1]?.text;

      console.log("Question:", question);
      console.log("Answer:", answerr);

    return answer || "Copilot did not respond";
     
    const afterSend = await client
      .api(`/beta/copilot/conversations/${this.conversationId}/chat`)
      .get();
      

    console.log("MESSAGES RIGHT AFTER SEND →", afterSend);

    // ⭐ STEP 4 — Poll for assistant response
      for (let i = 0; i < 12; i++) {

      await new Promise(resolve => setTimeout(resolve, 2000));

      const messages = await client
        .api(`/copilot/conversations/${this.conversationId}/chat`)
        .version("beta")
        .get();

      console.log("COPILOT RAW MESSAGES →", messages);

      const assistant = messages?.value
        ?.slice()
        ?.reverse()
        ?.find((m: any) =>
          m["@odata.type"] === "#microsoft.graph.copilotConversationResponseMessage"
        );

      if (assistant) {

        const text =
          assistant?.content
            ?.filter((c: any) => c.text)
            ?.map((c: any) => c.text)
            ?.join("\n") || "Empty Copilot reply";

        return text;
      }

    }

    console.warn("COPILOT TIMEOUT — no assistant message");
    return "Copilot did not respond (timeout)";
  }
}
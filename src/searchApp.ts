import {TeamsActivityHandler, CardFactory, TurnContext, MessagingExtensionQuery} from "botbuilder";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import {MessageExtensionTokenResponse, handleMessageExtensionQueryWithSSO, OnBehalfOfCredentialAuthConfig, OnBehalfOfUserCredential} from "@microsoft/teamsfx";
import { CommandIds } from "./enums/CommandIds";
import { EntityType } from "./enums/EntityType";
import "isomorphic-fetch";
import config from "./config";

const oboAuthConfig: OnBehalfOfCredentialAuthConfig = {
  authorityHost: config.authorityHost,
  clientId: config.clientId,
  tenantId: config.tenantId,
  clientSecret: config.clientSecret,
};
const initialLoginEndpoint = `https://${config.botDomain}/auth-start.html`;

export class SearchApp extends TeamsActivityHandler {
  constructor() {
    super();
  }

  public async handleTeamsMessagingExtensionQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<any> {

    return await handleMessageExtensionQueryWithSSO(context, oboAuthConfig, initialLoginEndpoint, ["User.Read.All", "User.Read"],
     async (token: MessageExtensionTokenResponse) => {
        const credential = new OnBehalfOfUserCredential(token.ssoToken, oboAuthConfig);
        const attachments = [];
        if (query.parameters[0] && query.parameters[0].name === "initialRun") 
        {
          // Return empty preview Items on initial run
          return this.GetPreviewItems(attachments);
        } 
        else 
        {
          // Create an instance of the TokenCredentialAuthenticationProvider by passing the tokenCredential instance and options to the constructor
          const authProvider = new TokenCredentialAuthenticationProvider(credential, 
            {scopes: ["User.Read.All", "Files.Read.All", "Calendars.Read", "People.Read", "Sites.Read.All", "Mail.Read"]});

          // Initialize Graph client instance with authProvider
          const graphClient = Client.initWithMiddleware({authProvider: authProvider});
          
          //const searchContext = query.parameters[0].value;
          let searchContext: string = (query.parameters[0]?.value as string) ?? '';

          //Get the entity type from the commandId
          const entityType = this.getEntityType(query.commandId);
          
          //Add the PromotedState:2 filter in order to get only news posts in Search API
          if (query.commandId === CommandIds.SearchNews) {
            searchContext = `${searchContext} PromotedState:2`;
          }

          const searchResponse = {requests: [{ entityTypes: [entityType],
            query: {queryString: searchContext},
            fields: ['Id','title','name','subject','webURL','start','createdDateTime','start','end']
           }]};
          const results = await graphClient.api('/search/query').post(searchResponse);

          if (results != null && results.value.length > 0)
          {
              const hitsContainer = results.value[0].hitsContainers[0];
              const total = hitsContainer.total;
              const moreResultsAvailable = hitsContainer.moreResultsAvailable;
              const hits = hitsContainer.hits;

              for (const item of hits) {
                const title = this.GetThumbnailCardTitle(item, entityType);
                const text = this.GetThumbnailCardText(item, entityType);

                const thumbnailCard = CardFactory.thumbnailCard(title, text);
                attachments.push(thumbnailCard);
              }
          }
        }
        return {
          composeExtension: {
            type: "result",
            attachmentLayout: "list",
            attachments: attachments,
          },
        };
      }
    );
  }

  private GetPreviewItems(attachments: any[]): any {
    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      }
    };
  }

  private getEntityType(commandId: string): EntityType {
    switch (commandId) {
      case CommandIds.SearchEvents:
        return EntityType.Event;
      case CommandIds.SearchFiles:
        return EntityType.DriveItem;
      case CommandIds.SearchListItems:
        return EntityType.ListItem;
      case CommandIds.SearchMessages:
        return EntityType.Message;
      case CommandIds.SearchNews:
        return EntityType.ListItem;
      default:
        return EntityType.UnknownFutureValue;
    }
  }

  private GetThumbnailCardTitle(item: any, entityType: EntityType): string {
    if (entityType === EntityType.Event) {
      return item.resource.subject || "Unknown";
    } else if (entityType === EntityType.ListItem) {
        return item.resource.fields.title || "Unknown";
    } else if (entityType === EntityType.Message) {
        return item.resource.subject || "Unknown";
    } else if (entityType === EntityType.DriveItem) {
        return item.resource.name || "Unknown";
    } else {
        return "Unknown";
    }
  }

  private GetThumbnailCardText(item: any, entityType: EntityType): string {
    if (entityType === EntityType.Event) {
        return `Start: ${item.resource.start.dateTime}` || "Unknown";
    } else if (entityType === EntityType.ListItem) {
        return `Created: ${item.resource.createdDateTime}` || "Unknown";
    } else if (entityType === EntityType.Message) {
        return `Created: ${item.resource.createdDateTime}` || "Unknown";
    } else if (entityType === EntityType.DriveItem) {
        return `Created: ${item.resource.createdDateTime}` || "Unknown";
    } else {
        return "Unknown";
    }
  }

  public async handleTeamsMessagingExtensionSelectItem(context: TurnContext, obj: any): Promise<any> {
    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [CardFactory.heroCard(obj.name, obj.description)],
      },
    };
  }
}

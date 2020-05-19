import {
  AuthenticationProvider,
  Client,
  ClientOptions
} from "@microsoft/microsoft-graph-client";

/**
 * Authentication provider which uses client credentials
 * Basically this is a port of the dotnet implementation.
 */
export class ClientCredentialProvider implements AuthenticationProvider {
  private clientId: string;

  private secret: string;

  private tenantId: string;

  private accessToken: string;

  constructor(clientId: string, secret: string, tenantId: string) {
    this.clientId = clientId;
    this.secret = secret;
    this.tenantId = tenantId;
  }

  async getAccessToken(): Promise<string> {
    if (!this.accessToken) {
      const accessToken = await ClientCredentialProvider.callTokenApi(
        this.clientId,
        this.secret,
        this.tenantId
      );
      // TODO check if accessToken is expired, and refresh if required
      this.accessToken = accessToken;
    }
    return this.accessToken;
  }

  /**
   * Call the token API to obtain an access token
   */
  static async callTokenApi(
    clientId: string,
    secret: string,
    tenantId: string
  ): Promise<string> {
    // Build the request
    const params = new URLSearchParams();
    params.append("client_id", clientId);
    params.append("client_secret", secret);
    params.append("scope", "https://graph.microsoft.com/.default");
    params.append("grant_type", "client_credentials");

    const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const postRequest = new Request(url, {
      method: "POST",
      headers: new Headers({
        "Cache-control": "no-cache",
        "Content-Type": "application/x-www-form-urlencoded"
      }),
      body: params.toString()
    });

    // call the token endpoint
    const response = await fetch(postRequest).catch(error => {
      throw new Error(`Error trying to obtain access token: ${error}`);
    });

    // throw error if its not ok
    if (!response.ok) {
      throw new Error(`Failed to obtain access token: ${response.status}`);
    }

    // parse the response
    const responseJson = await response.json();

    // return the access token
    return responseJson.access_token;
  }
}

/**
 * Get a Graph API Client which is uses app access token
 */
export const getGraphClient = async (
  clientId: string,
  secret: string,
  tenantId: string
): Promise<Client> => {
  // create a client here
  const clientOptions: ClientOptions = {
    authProvider: new ClientCredentialProvider(clientId, secret, tenantId)
  };
  return Client.initWithMiddleware(clientOptions);
};

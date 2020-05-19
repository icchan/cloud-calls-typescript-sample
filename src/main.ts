import "isomorphic-fetch";
import {
  InvitationParticipantInfo,
  Call
} from "@microsoft/microsoft-graph-types";
import { getGraphClient } from "./calls_api/calls_service";

const clientId = "ab49ad75-e3df-4481-8048-790fc6f77537";
const secret = "isqu3pO/Nnut6XXGN4Iv4AAt==OLwvH[";
const tenantId = "1ada90b0-df6e-47e9-8150-c662cb71fc04";
const userids = [
  "be819f46-3a31-4b26-adbd-ce39a74d33df",
  "e66f2484-441e-43d6-95f8-ce9009940832"
];

const callback = "";

(async (): Promise<void> => {
  try {
    const client = await getGraphClient(clientId, secret, tenantId);

    const targetUsers: InvitationParticipantInfo[] = userids.map(function(
      userid
    ) {
      return {
        identity: {
          user: {
            id: userid
          }
        }
      };
    });

    // create a call request
    const call: Call = {
      callbackUri: callback,
      source: {
        identity: {
          application: {
            displayName: "My Test App",
            id: clientId
          }
        }
      },
      targets: targetUsers,
      requestedModalities: ["audio"],
      mediaConfig: {
        "@odata.type": "#microsoft.graph.serviceHostedMediaConfig"
      },
      tenantId
    };

    // invoked the calls api
    const response = await client.api("/communications/calls").post(call);

    console.log(JSON.stringify(response, null, 4));
  } catch (e) {
    console.log(`Sorry, something went bad ${JSON.stringify(e)}`);
  }
})();

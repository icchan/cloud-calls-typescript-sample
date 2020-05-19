import { Call } from "@microsoft/microsoft-graph-types";

// This is essential infromation for joining existing online meeting
export type existingMeetingInfo = {
  threadId: string;
  messageId: string;
  tenantId: string;
  organizerId: string;
};

export const getMeetingInfo = (joinUrl: string): existingMeetingInfo => {
  /*
  joinUrl has meeting information such as thread id, message id, tenand id and organizer id
  When we split it with '/', we can get string array.
    0: "https:"
    1: ""
    2: "teams.microsoft.com"
    3: "l"
    4: "meetup-join"
    5: "19:meeting_NjNiNzRlODYtNTllYi00NzNkLWE4NzYtOTIzMmFmNThmMmEx@thread.v2"
    6: "0?context={
        "Tid":"b21a0d16-4e90-4cdb-a05b-ad3846369881"
        ,"Oid":"ea7140cd-bced-4bdf-931b-06cc30891bb8"}"

    Index 5 includes thread id

    Index 6 includes message id, tenant id and organizer id.
    We can split value of index 6 with '?contest' then we can get following string array
    0: "0"
    1: "{"Tid":"b21a0d16-4e90-4cdb-a05b-ad3846369881","Oid":"ea7140cd-bced-4bdf-931b-06cc30891bb8"}"

    In this case, message id is 0, tenant id is value of Tid, organizer id is value of Oid
   */
  const THREAD_ID_INDEX = 5;
  const MEETING_INFO_INDEX = 6;
  const TENANT_AND_ORGANIZER_INDEX = 1;
  const MESSAGE_ID_INDEX = 0;

  const decodedUri: string[] = decodeURIComponent(joinUrl).split("/");
  const meetingInfo: string[] = decodedUri[MEETING_INFO_INDEX].split(
    "?context="
  );
  const organizerInfo: {
    Tid: string;
    Oid: string;
  } = JSON.parse(meetingInfo[TENANT_AND_ORGANIZER_INDEX]);

  return {
    threadId: decodedUri[THREAD_ID_INDEX],
    messageId: meetingInfo[MESSAGE_ID_INDEX],
    tenantId: organizerInfo.Tid,
    organizerId: organizerInfo.Oid
  };
};

export const createJoinCallRequest = (
  meetingInfo: existingMeetingInfo,
  callbackUri: string
): Call => {
  const call: Call = {
    callbackUri,
    requestedModalities: ["audio"],
    tenantId: meetingInfo.tenantId,
    mediaConfig: {
      "@odata.type": "#microsoft.graph.serviceHostedMediaConfig"
    },
    chatInfo: {
      threadId: meetingInfo.threadId,
      messageId: meetingInfo.messageId
    },
    meetingInfo: {
      "@odata.type": "#microsoft.graph.organizerMeetingInfo",
      organizer: {
        "@odata.type": "#microsoft.graph.identitySet",
        user: {
          "@odata.type": "#microsoft.graph.identity",
          id: meetingInfo.organizerId,
          tenantId: meetingInfo.tenantId
        }
      },
      allowConversationWithoutHost: true
    }
  };

  return call;
};

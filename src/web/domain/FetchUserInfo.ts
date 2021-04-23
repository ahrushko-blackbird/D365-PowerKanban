import * as WebApiClient from "xrm-webapi-client";

const avatarCache: {[key: string]: Promise<string>} = {};

export const FetchUserAvatar = async (id: string) => {
    if (avatarCache[id]) {
        return Promise.resolve(avatarCache[id]);
    }

    avatarCache[id] = WebApiClient.Retrieve({ entityName: "systemuser", entityId: id, queryParams: "?$select=entityimage_url" })
    .then((data: any) => data.entityimage_url);

    return avatarCache[id];
};
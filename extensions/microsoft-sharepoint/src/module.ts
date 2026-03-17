import { createExtension } from "@cognigy/extension-tools";

import { getSharepointSiteInfoNode } from "./nodes/getSharepointSiteInfo";
import { cloudConnection } from "./connections/cloudConnection";
import { getSharepointListItemsNode } from "./nodes/getSharepointListItems";
import { basicConnection } from "./connections/basicConnection";
import { graphConnection } from "./connections/graphConnection";
import { sharepointKnowledgeConnector } from "./knowledge-connectors/sharepointKnowledgeConnector";


export default createExtension({
	nodes: [
		getSharepointSiteInfoNode,
		getSharepointListItemsNode
	],

	connections: [
		cloudConnection,
		basicConnection,
		graphConnection
	],

	knowledge: [
		sharepointKnowledgeConnector
	],

	options: {
		label: "Microsoft Sharepoint"
	}
});
import { DateTime } from 'luxon';
import {
	type IPollFunctions,
	type IDataObject,
	type INodeExecutionData,
	type INodeType,
	type INodeTypeDescription,
	NodeConnectionTypes,
} from 'n8n-workflow';

import { getPath, microsoftApiRequest, microsoftApiRequestAllItemsDelta } from './GenericFunctions';
import { triggerDescription } from './TriggerDescription';

export class MicrosoftOneDriveWithDriveIdTrigger implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Microsoft OneDrive Trigger (With Drive ID)',
		name: 'microsoftOneDriveWithDriveIdTrigger',
		icon: 'file:oneDrive.svg',
		group: ['trigger'],
		version: 1,
		description: 'Trigger for Microsoft OneDrive API with custom Drive ID.',
		subtitle: '={{($parameter["event"])}}',
		defaults: {
			name: 'Microsoft OneDrive Trigger (With Drive ID)',
		},
		credentials: [
			{
				name: 'microsoftOneDriveWithDriveIdOAuth2Api',
				required: true,
			},
		],
		polling: true,
		inputs: [],
		outputs: [NodeConnectionTypes.Main],
		properties: [...triggerDescription],
	};

	methods = {
		loadOptions: {},
	};

	async poll(this: IPollFunctions): Promise<INodeExecutionData[][] | null> {
		const workflowData = this.getWorkflowStaticData('node');
		let responseData: IDataObject[];

		const credentials = await this.getCredentials('microsoftOneDriveWithDriveIdOAuth2Api');
		const baseUrl = (
			typeof credentials.graphApiBaseUrl === 'string' && credentials.graphApiBaseUrl !== ''
				? credentials.graphApiBaseUrl
				: 'https://graph.microsoft.com'
		).replace(/\/+$/, '');

		const driveId =
			typeof credentials.driveId === 'string' && credentials.driveId !== ''
				? credentials.driveId
				: undefined;

		const lastLink: string =
			(workflowData.LastLink as string) ||
			`${baseUrl}/v1.0/drives/${driveId}/root/delta?token=latest`;

		const now = DateTime.now().toUTC();
		const start = DateTime.fromISO(workflowData.lastTimeChecked as string) || now;
		const end = now;
		const event = this.getNodeParameter('event', 'fileCreated') as string;
		const watch = this.getNodeParameter('watch', 'anyFile') as string;
		const watchFolder = (this.getNodeParameter('watchFolder', false) as boolean) || false;
		const folderChild = (this.getNodeParameter('options.folderChild', false) as boolean) || false;

		let eventType = 'created';
		let eventResource = 'file';
		if (event.includes('Updated')) {
			eventType = 'updated';
		}
		if (event.includes('folder')) {
			eventResource = 'folder';
		}
		try {
			if (this.getMode() === 'manual') {
				const endpoint = driveId
					? `${baseUrl}/v1.0/drives/${driveId}/root/delta`
					: `${baseUrl}/v1.0/me/drive/root/delta`;
				responseData = (
					await microsoftApiRequest.call(this, 'GET', '', {}, {}, endpoint)
				).value as IDataObject[];
			} else {
				const response: IDataObject = (await microsoftApiRequestAllItemsDelta.call(
					this,
					lastLink,
					start,
					eventType,
				)) as IDataObject;
				responseData = response.returnData as IDataObject[];
				workflowData.LastLink = response.deltaLink;
			}

			workflowData.lastTimeChecked = end.toISO();
			if (watch === 'selectedFile') {
				const fileId = (
					this.getNodeParameter('fileId', '', {
						extractValue: true,
					}) as string
				).replace('%21', '!');
				if (fileId) {
					responseData = responseData.filter((item: IDataObject) => item.id === fileId);
				}
			}

			if (
				!folderChild &&
				(watch === 'oneSelectedFolder' || watch === 'selectedFolder' || watchFolder)
			) {
				const folderId = (
					this.getNodeParameter('folderId', '', {
						extractValue: true,
					}) as string
				).replace('%21', '!');
				if (folderId) {
					if (watch === 'oneSelectedFolder') {
						responseData = responseData.filter((item: IDataObject) => item.id === folderId);
					} else {
						responseData = responseData.filter(
							(item: IDataObject) => (item.parentReference as IDataObject).id === folderId,
						);
					}
				}
			}
			if (folderChild && (watch === 'selectedFolder' || watchFolder)) {
				const folderId = (
					this.getNodeParameter('folderId', '', {
						extractValue: true,
					}) as string
				).replace('%21', '!');
				const folderPath = await getPath.call(this, folderId);

				responseData = responseData.filter((item: IDataObject) => {
					const path = (item.parentReference as IDataObject)?.path as string;
					return typeof path === 'string' && path.startsWith(folderPath);
				});
			}
			responseData = responseData.filter((item: IDataObject) => item[eventResource]);
			if (!responseData?.length) {
				return null;
			}

			const simplify = this.getNodeParameter('simple') as boolean;
			if (simplify) {
				responseData = responseData.map((x) => ({
					id: x.id,
					createdDateTime: (x.fileSystemInfo as IDataObject)?.createdDateTime,
					lastModifiedDateTime: (x.fileSystemInfo as IDataObject)?.lastModifiedDateTime,
					name: x.name,
					webUrl: x.webUrl,
					size: x.size,
					path: (x.parentReference as IDataObject)?.path || '',
					mimeType: (x.file as IDataObject)?.mimeType || '',
				}));
			}

			return [this.helpers.returnJsonArray(responseData)];
		} catch (error) {
			if (this.getMode() === 'manual' || !workflowData.lastTimeChecked) {
				throw error;
			}
			const workflow = this.getWorkflow();
			const node = this.getNode();
			const errorObj = error as Error & { description?: string };
			this.logger.error(
				`There was a problem in '${node.name}' node in workflow '${workflow.id}': '${errorObj.description}'`,
				{
					node: node.name,
					workflowId: workflow.id,
					error,
				},
			);
			throw error;
		}

		return null;
	}
}
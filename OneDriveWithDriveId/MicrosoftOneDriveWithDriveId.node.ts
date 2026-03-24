import { IncomingMessage } from 'http';
import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	JsonObject,
} from 'n8n-workflow';
import { NodeApiError, NodeConnectionTypes, NodeOperationError } from 'n8n-workflow';

import { fileFields, fileOperations } from './FileDescription';
import { folderFields, folderOperations } from './FolderDescription';
import { microsoftApiRequest, microsoftApiRequestAllItems } from './GenericFunctions';

export class MicrosoftOneDriveWithDriveId implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Microsoft OneDrive (With Drive ID)',
		name: 'microsoftOneDriveWithDriveId',
		icon: 'file:oneDrive.svg',
		group: ['input'],
		version: [1, 2],
		subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
		description: 'Consume Microsoft OneDrive API with custom Drive ID',
		schemaPath: 'Microsoft/OneDrive',
		defaults: {
			name: 'Microsoft OneDrive (With Drive ID)',
		},
		usableAsTool: true,
		inputs: [NodeConnectionTypes.Main],
		outputs: [NodeConnectionTypes.Main],
		credentials: [
			{
				name: 'microsoftOneDriveWithDriveIdOAuth2Api',
				required: true,
			},
		],
		properties: [
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'File',
						value: 'file',
					},
					{
						name: 'Folder',
						value: 'folder',
					},
				],
				default: 'file',
			},
			...fileOperations,
			...fileFields,
			...folderOperations,
			...folderFields,
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];
		const length = items.length;
		const nodeVersion = this.getNode().typeVersion;
		let responseData;
		const resource = this.getNodeParameter('resource', 0);
		const operation = this.getNodeParameter('operation', 0);
		for (let i = 0; i < length; i++) {
			try {
				if (resource === 'file') {
					if (operation === 'copy') {
						const fileId = this.getNodeParameter('fileId', i) as string;
						const additionalFields = this.getNodeParameter('additionalFields', i);
						const parentReference = this.getNodeParameter('parentReference', i) as IDataObject;
						const body: IDataObject = {};
						if (parentReference) {
							body.parentReference = { ...parentReference, driveId: undefined };
						}
						if (additionalFields.name) {
							body.name = additionalFields.name as string;
						}
						responseData = await microsoftApiRequest.call(
							this,
							'POST',
							`/drive/items/${fileId}/copy`,
							body,
							{},
							undefined,
							{},
							{ json: true, resolveWithFullResponse: true },
						);
						responseData = { location: responseData.headers.location };
					}
					if (operation === 'delete') {
						const fileId = this.getNodeParameter('fileId', i) as string;
						responseData = await microsoftApiRequest.call(this, 'DELETE', `/drive/items/${fileId}`);
						responseData = { success: true };
					}
					if (operation === 'download') {
						const fileId = this.getNodeParameter('fileId', i) as string;
						const dataPropertyNameDownload = this.getNodeParameter('binaryPropertyName', i);
						responseData = await microsoftApiRequest.call(this, 'GET', `/drive/items/${fileId}`);

						const fileName = responseData.name;
						const downloadUrl = responseData['@microsoft.graph.downloadUrl'];

						if (responseData.file === undefined) {
							throw new NodeApiError(this.getNode(), responseData as JsonObject, {
								message: 'The ID you provided does not belong to a file.',
							});
						}

						let mimeType: string | undefined;
						if (responseData.file.mimeType) {
							mimeType = responseData.file.mimeType;
						}

						try {
							responseData = await microsoftApiRequest.call(
								this,
								'GET',
								`/drive/items/${fileId}/content`,
								{},
								{},
								undefined,
								{},
								{ encoding: null, resolveWithFullResponse: true },
							);
						} catch (error) {
							if (downloadUrl) {
								try {
									responseData = await this.helpers.httpRequest({
										method: 'GET',
										url: downloadUrl,
										returnFullResponse: true,
										encoding: 'arraybuffer',
										json: false,
									});
								} catch (downloadError) {
									throw new NodeApiError(this.getNode(), downloadError as JsonObject, {
										message: 'Failed to download file from both primary and fallback URLs',
									});
								}
							} else {
								throw new NodeApiError(this.getNode(), error as JsonObject);
							}
						}

						const newItem: INodeExecutionData = {
							json: items[i].json,
							binary: {},
						};

						if (mimeType === undefined && responseData.headers['content-type']) {
							mimeType = responseData.headers['content-type'];
						}

						if (items[i].binary !== undefined) {
							Object.assign(newItem.binary!, items[i].binary);
						}

						items[i] = newItem;

						let data;
						if (responseData?.body instanceof IncomingMessage) {
							data = responseData.body;
						} else {
							data = Buffer.from(responseData.body as Buffer);
						}

						items[i].binary![dataPropertyNameDownload] = await this.helpers.prepareBinaryData(
							data,
							fileName as string,
							mimeType,
						);
					}
					if (operation === 'get') {
						const fileId = this.getNodeParameter('fileId', i) as string;
						responseData = await microsoftApiRequest.call(this, 'GET', `/drive/items/${fileId}`);
					}
					if (operation === 'search') {
						const query = this.getNodeParameter('query', i) as string;
						responseData = await microsoftApiRequestAllItems.call(
							this,
							'value',
							'GET',
							`/drive/root/search(q='${encodeURIComponent(query)}')`,
						);
						responseData = responseData.filter((item: IDataObject) => item.file);
					}
					if (operation === 'share') {
						const fileId = this.getNodeParameter('fileId', i) as string;
						const type = this.getNodeParameter('type', i) as string;
						const scope = this.getNodeParameter('scope', i) as string;
						const body: IDataObject = {
							type,
							scope,
						};
						responseData = await microsoftApiRequest.call(
							this,
							'POST',
							`/drive/items/${fileId}/createLink`,
							body,
						);
					}
					if (operation === 'upload') {
						const parentId = this.getNodeParameter('parentId', i) as string;
						const isBinaryData = this.getNodeParameter('binaryData', i);
						const fileName = this.getNodeParameter('fileName', i) as string;

						if (isBinaryData) {
							const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i);
							const binaryData = this.helpers.assertBinaryData(i, binaryPropertyName);
							const body = await this.helpers.getBinaryDataBuffer(i, binaryPropertyName);
							let encodedFilename;

							if (nodeVersion >= 1.1) {
								if (fileName !== '') {
									encodedFilename = encodeURIComponent(fileName);
								} else if (binaryData.fileName !== undefined) {
									encodedFilename = encodeURIComponent(binaryData.fileName);
								}
							} else {
								if (fileName !== '') {
									encodedFilename = encodeURIComponent(fileName);
								}

								if (binaryData.fileName !== undefined) {
									encodedFilename = encodeURIComponent(binaryData.fileName);
								}
							}

							responseData = await microsoftApiRequest.call(
								this,
								'PUT',
								`/drive/items/${parentId}:/${encodedFilename}:/content`,
								body,
								{},
								undefined,
								{ 'Content-Type': binaryData.mimeType, 'Content-Length': body.length },
								{},
							);

							responseData = JSON.parse(responseData as string);
						} else {
							const body = this.getNodeParameter('fileContent', i) as string;
							if (fileName === '') {
								throw new NodeOperationError(this.getNode(), 'File name must be set!', {
									itemIndex: i,
								});
							}
							const encodedFilename = encodeURIComponent(fileName);
							responseData = await microsoftApiRequest.call(
								this,
								'PUT',
								`/drive/items/${parentId}:/${encodedFilename}:/content`,
								body,
								{},
								undefined,
								{ 'Content-Type': 'text/plain' },
							);
						}
					}
				}
				if (resource === 'folder') {
					if (operation === 'create') {
						const names = (this.getNodeParameter('name', i) as string)
							.split('/')
							.filter((s) => s.trim() !== '');
						const options = this.getNodeParameter('options', i);
						let parentFolderId = options.parentFolderId ? options.parentFolderId : null;
						for (const name of names) {
							const body: IDataObject = {
								name,
								folder: {},
							};
							let endpoint = '/drive/root/children';
							if (parentFolderId) {
								endpoint = `/drive/items/${parentFolderId}/children`;
							}
							responseData = await microsoftApiRequest.call(this, 'POST', endpoint, body);
							if (!responseData.id) {
								break;
							}
							parentFolderId = responseData.id;
						}
					}
					if (operation === 'delete') {
						const folderId = this.getNodeParameter('folderId', i) as string;
						responseData = await microsoftApiRequest.call(
							this,
							'DELETE',
							`/drive/items/${folderId}`,
						);
						responseData = { success: true };
					}
					if (operation === 'getChildren') {
						const folderId = this.getNodeParameter('folderId', i) as string;
						responseData = await microsoftApiRequestAllItems.call(
							this,
							'value',
							'GET',
							`/drive/items/${folderId}/children`,
						);
					}
					if (operation === 'search') {
						const query = this.getNodeParameter('query', i) as string;
						responseData = await microsoftApiRequestAllItems.call(
							this,
							'value',
							'GET',
							`/drive/root/search(q='${encodeURIComponent(query)}')`,
						);
						responseData = responseData.filter((item: IDataObject) => item.folder);
					}
					if (operation === 'share') {
						const folderId = this.getNodeParameter('folderId', i) as string;
						const type = this.getNodeParameter('type', i) as string;
						const scope = this.getNodeParameter('scope', i) as string;
						const body: IDataObject = {
							type,
							scope,
						};
						responseData = await microsoftApiRequest.call(
							this,
							'POST',
							`/drive/items/${folderId}/createLink`,
							body,
						);
					}
				}
				if (resource === 'file' || resource === 'folder') {
					if (operation === 'rename') {
						const itemId = this.getNodeParameter('itemId', i) as string;
						const newName = this.getNodeParameter('newName', i) as string;
						const body = { name: newName };
						responseData = await microsoftApiRequest.call(
							this,
							'PATCH',
							`/drive/items/${itemId}`,
							body,
						);
					}
				}
			} catch (error) {
				if (this.continueOnFail()) {
					if (resource === 'file' && operation === 'download') {
						items[i].json = { error: (error as Error).message };
					} else {
						const executionErrorData = this.helpers.constructExecutionMetaData(
							this.helpers.returnJsonArray({ error: (error as Error).message }),
							{ itemData: { item: i } },
						);
						returnData.push(...executionErrorData);
					}
					continue;
				}
				throw error;
			}
			if (responseData === undefined) {
				throw new NodeOperationError(this.getNode(), 'Operation not implemented', { itemIndex: i });
			}
			const executionData = this.helpers.constructExecutionMetaData(
				this.helpers.returnJsonArray(responseData as IDataObject),
				{ itemData: { item: i } },
			);

			returnData.push(...executionData);
		}
		if (resource === 'file' && operation === 'download') {
			return [items];
		}

		return [returnData];
	}
}
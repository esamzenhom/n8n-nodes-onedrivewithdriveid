import { type IPollFunctions, type INodeExecutionData, type INodeType, type INodeTypeDescription } from 'n8n-workflow';
export declare class MicrosoftOneDriveWithDriveIdTrigger implements INodeType {
    description: INodeTypeDescription;
    methods: {
        loadOptions: {};
    };
    poll(this: IPollFunctions): Promise<INodeExecutionData[][] | null>;
}
//# sourceMappingURL=MicrosoftOneDriveWithDriveIdTrigger.node.d.ts.map
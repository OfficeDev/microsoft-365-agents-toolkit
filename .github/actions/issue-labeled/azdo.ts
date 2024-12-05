import * as vm from 'azure-devops-node-api';
import * as nodeApi from 'azure-devops-node-api';
import * as WorkItemTrackingApi from 'azure-devops-node-api/WorkItemTrackingApi';
import * as WorkItemTrackingInterfaces from 'azure-devops-node-api/interfaces/WorkItemTrackingInterfaces';
import {
	JsonPatchDocument,
	JsonPatchOperation,
	Operation,
} from 'azure-devops-node-api/interfaces/common/VSSInterfaces';

class ItemInfo {
	url: string;
	id: number;
	public constructor(url: string, id: number) {
		this.url = url;
		this.id = id;
	}
}

export class DevopsClient {
	token: string;
	org: string;
	projectId: string;
	bugArea: string;
	bugIteration: string;
	featureArea: string;
	featureIteration: string;

	witApi?: WorkItemTrackingApi.IWorkItemTrackingApi;

	constructor(
		token: string,
		org: string,
		projectId: string,
		bugArea: string,
		bugIteration: string,
		featureArea: string,
		featureIteration: string,
	) {
		this.token = token;
		this.org = org;
		this.projectId = projectId;
		this.bugArea = bugArea;
		this.bugIteration = bugIteration;
		this.featureArea = featureArea;
		this.featureIteration = featureIteration;
	}

	public async init() {
		let orgUrl = `https://dev.azure.com/${this.org}`;
		const webApi: nodeApi.WebApi = await this.getApi(orgUrl);
		this.witApi = await webApi.getWorkItemTrackingApi();
	}

	public async queryPreviousItem(description: string): Promise<ItemInfo | undefined> {
		var query = `Select [System.Id] From WorkItems Where [System.Description] Contains Words '${description}' AND [System.HyperLinkCount] > 0 AND [State] <> 'Removed' order by [Microsoft.VSTS.Common.Priority] asc, [System.CreatedDate] desc`;
		const items = await this.witApi!.queryByWiql({ query: query }, undefined, undefined, 1);
		if (items.workItems?.length && items.workItems[0].id) {
			const resp = await this.witApi?.getWorkItem(items.workItems[0].id);
			return {
				url: resp?._links?.html?.href,
				id: items.workItems[0].id,
			};
		} else {
			return undefined;
		}
	}

	public async createFeatureItem(
		titleValue: string,
		asigneeValue: string | undefined,
		tagsValue: string | undefined,
		url: string,
		sprintPath: string,
	): Promise<WorkItemTrackingInterfaces.WorkItem> {
		return this.createItem(
			titleValue,
			asigneeValue,
			this.featureArea,
			this.featureIteration,
			tagsValue,
			url,
			'Feature',
			sprintPath,
		);
	}

	public async createBugItem(
		titleValue: string,
		asigneeValue: string | undefined,
		tagsValue: string | undefined,
		url: string,
		sprintPath: string,
	): Promise<WorkItemTrackingInterfaces.WorkItem> {
		return this.createItem(
			titleValue,
			asigneeValue,
			this.bugArea,
			this.bugIteration,
			tagsValue,
			url,
			'Bug',
			this.bugIteration,
		);
	}

	public async createItem(
		titleValue: string,
		asigneeValue: string | undefined,
		areaValue: string,
		iterationValue: string,
		tagsValue: string | undefined,
		url: string,
		type: string,
		sprintPath: string,
	): Promise<WorkItemTrackingInterfaces.WorkItem> {
		let document: JsonPatchOperation[] = [];

		const title: JsonPatchOperation = {
			path: '/fields/System.Title',
			op: Operation.Add,
			value: titleValue,
		};
		document.push(title);

		if (asigneeValue) {
			const asignee: JsonPatchOperation = {
				path: '/fields/System.AssignedTo',
				op: Operation.Add,
				value: asigneeValue,
			};
			document.push(asignee);
		}

		const area: JsonPatchOperation = {
			path: '/fields/System.AreaPath',
			op: Operation.Add,
			value: areaValue,
		};
		document.push(area);

		const iteration: JsonPatchOperation = {
			path: '/fields/System.IterationPath',
			op: Operation.Add,
			value: sprintPath ?? iterationValue,
		};
		document.push(iteration);

		if (tagsValue) {
			const tags: JsonPatchOperation = {
				path: '/fields/System.Tags',
				op: Operation.Add,
				value: tagsValue,
			};
			document.push(tags);
		}

		const previous = await this.queryPreviousItem(url);
		const description: JsonPatchOperation = {
			path: '/fields/System.Description',
			op: Operation.Add,
			value: this.buildDescription(url, previous),
		};
		document.push(description);

		const hyperLink: JsonPatchOperation = {
			path: '/relations/-',
			op: Operation.Add,
			value: {
				rel: 'Hyperlink',
				url: url,
				attributes: { comment: 'github issue link' },
			},
		};
		document.push(hyperLink);

		const item = await this.witApi!.createWorkItem(
			undefined,
			document as JsonPatchDocument,
			this.projectId,
			type,
		);
		return item;
	}

	private async getApi(serverUrl: string): Promise<vm.WebApi> {
		let authHandler = vm.getHandlerFromToken(this.token);
		let vsts: vm.WebApi = new vm.WebApi(serverUrl, authHandler);
		await vsts.connect();
		return vsts;
	}

	private buildDescription(url: string, addition?: ItemInfo): string {
		var description = `<a href="${url}">${url}</a>`;
		if (addition) {
			description += `<br>There is an existing work item related to this issue<br>`;
			description += `<div><a href="${addition.url}" data-vss-mention="version:1.0">#${addition.id}</a>&nbsp;</div>`;
		}
		return description;
	}
}

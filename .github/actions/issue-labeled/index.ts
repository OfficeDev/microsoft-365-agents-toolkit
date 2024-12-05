import { OctoKit, OctoKitIssue } from '../api/octokit';
import { Action } from '../common/Action';
import { DevopsClient } from './azdo';
import { getRequiredInput, safeLog } from '../common/utils';
import { context } from '@actions/github';
import { getInput } from '@actions/core';
import { getEmail, sendAlert } from '../teamsfx-utils/utils';
import * as WorkItemTrackingInterfaces from 'azure-devops-node-api/interfaces/WorkItemTrackingInterfaces';
import { AzureCliCredential } from "@azure/identity";


const githubToken = getRequiredInput('token');
const org = getRequiredInput('devops-org');
const projectId = getRequiredInput('devops-projectId');
const titlePreix = getRequiredInput('title-prefix');
const bugLabel = getRequiredInput('bug-label');
const bugArea = getRequiredInput('bug-area-path');
const bugIteration = getRequiredInput('bug-iteration-path');
const featureArea = getInput('feature-area-path');
const featureIteration = getInput('feature-iteration-path');

class Labeled extends Action {
	id = 'Labeled';

	async onLabeled(issue: OctoKitIssue) {
		const content = await issue.getIssue();
		let sprintPath = "";
		if (content.milestone?.description) {
			const match = content.milestone?.description.match(/Sprint path is:(.*)/);
			if (match && match.length > 1) {
				sprintPath = match[1];
			}
		}
		safeLog(`the issue ${content.number} is created by label`);
		let client = await this.createClient();
		const asignee = getEmail(content.assignee);
		if (!asignee) {
			safeLog(`the issue ${content.number} assignee:${content.assignee} is not associated with email address, ignore.`);
			const subject = '[Github Issue Alert] missing associated email address for assignee';
			const issueLink = `https://github.com/OfficeDev/TeamsFx/issues/${content.number}`;
			const fileLink = "https://github.com/OfficeDev/TeamsFx/blob/dev/.github/accounts.json";
			const message = `There is a github issue <a>${issueLink}</a> labeled with account <b>${content.assignee}</b> which is not associated with company email. Please check it and update the account mapping in the file <a>${fileLink}</a>.`;
			safeLog(message);
			sendAlert(subject, message);
		}
		const url = this.issueUrl(content.number);
		const title = titlePreix + content.title;
		let workItem: WorkItemTrackingInterfaces.WorkItem;
		safeLog(`issue labeled with ${bugLabel}. Bug work item will be created. ${sprintPath}`);
		workItem = await client.createBugItem(title, asignee, undefined, url, sprintPath);
		safeLog(`finished to create work item.`);
		const workItemUrl = workItem._links?.html?.href;
		if (workItemUrl) {
			await issue.postComment(`The issue is labeled with sprint and a work item created: ${workItemUrl}`);
		} else {
			safeLog(`no work item url found, ignore to post comment.`);
		}
	}
	async onTriggered(_: OctoKit) {
		const issueNumber = process.env.ISSUE_NUMBER;
		safeLog(`start manually create work item for issue ${issueNumber}`);
		const issue = new OctoKitIssue(githubToken, context.repo, { number: parseInt(issueNumber || "0") });
		await this.onLabeled(issue);
	}

	private async createClient() {
		let credential = new AzureCliCredential();
		const devopsToken = await credential.getToken("https://app.vssps.visualstudio.com/.default");

		let client = new DevopsClient(
			devopsToken.token,
			org,
			projectId,
			bugArea,
			bugIteration,
			featureArea,
			featureIteration,
		);
		await client.init();
		return client;
	}

	private issueUrl(id: number) {
		return `https://github.com/${context.repo.owner}/${context.repo.repo}/issues/${id}`;
	}
}

new Labeled().run(); // eslint-disable-line

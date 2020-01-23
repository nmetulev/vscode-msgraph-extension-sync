// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import * as adal from 'adal-node';

import "isomorphic-fetch";
import { Client, AuthenticationProvider, AuthenticationProviderOptions } from "@microsoft/microsoft-graph-client";

class Provider implements AuthenticationProvider {

	private token: string;

	constructor(token: string) {
		this.token = token;
	}

	public async getAccessToken(authenticationProviderOptions?: AuthenticationProviderOptions): Promise<string> {
		return this.token;
	}
}

// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {

	// Use the console to output diagnostic information (console.log) and errors (console.error)
	// This line of code will only be executed once when your extension is activated
	console.log('Congratulations, your extension "msgraph-extentions-sync" is now active!');

	// The command has been defined in the package.json file
	// Now provide the implementation of the command with registerCommand
	// The commandId parameter must match the command field in package.json
	let disposable = vscode.commands.registerCommand('extension.helloWorld', () => {
		// The code you place here will be executed every time your command is executed

		const extensions = vscode.extensions.all.filter(ext => !ext.packageJSON.isBuiltin).map(ext => ext.id);

		let resource = 'https://graph.microsoft.com';
		let clientId = '70fa51d4-dc5a-47c9-969a-487f825092dc';
		let authority = 'https://login.microsoftonline.com/common';

		let context = new adal.AuthenticationContext(authority, false);
		context.acquireUserCode(resource, clientId, 'en-us', (error, response) => {
			if (error) {
				console.log(error);
			} else {
				console.log('success getting user code', response);
				vscode.env.clipboard.writeText(response.userCode).then(() => {
					vscode.env.openExternal(vscode.Uri.parse(response.verificationUrl));
					context.acquireTokenWithDeviceCode(resource, clientId, response, async (error, response) => {
						if (error) {
							console.log(error);
						} else {
							console.log('success getting token', response);

							const options = {
								authProvider: new Provider((response as adal.TokenResponse).accessToken)
							};
							const client = Client.initWithMiddleware(options);

							try {
								// let result = await client.api('https://graph.microsoft.com/v1.0/me/extensions').post({
								// 	"@odata.type": "microsoft.graph.openTypeExtension",
								// 	"extensionName": "metulev.msgraph.extensions.sync",
								// 	extensions
								// });

								let result = await client
									.api('https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,mobilePhone&$expand=extensions')
									.get();
	
								console.log('extensions: ', result);
							} catch (ex) {
								console.log(ex);
							}
							

							// vscode.commands.executeCommand(
							// 	"workbench.extensions.installExtension",
							// 	name
							//   );


							

							

							

						}
					});
				});
			}
		});

		// Display a message box to the user
		vscode.window.showInformationMessage('Hello World!');
	});

	context.subscriptions.push(disposable);
}

// this method is called when your extension is deactivated
export function deactivate() {}

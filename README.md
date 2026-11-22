# SmartMD

SmartMD is an Obsidian plugin for publishing Markdown, syncing notes, and reviewing AI edits before applying them. The plugin interface is in Simplified Chinese. This release requires Obsidian 1.15.1 or later.

## Getting started

- **AI editing.** In SmartMD settings, configure an OpenAI-compatible endpoint and model. Right-click a note or selection to open AI editing. You can review the suggested changes before applying them and undo an applied edit.
- **SmartKeep publishing.** Sign in from SmartMD, then select the note or folder to publish. Choose public or password-protected access. If you enable continued publishing, changes within the published scope can be sent to SmartKeep.
- **OneDrive sync (Windows and Android).** Sign in to a Microsoft account on each device. On Android, use an Obsidian vault created with **App storage**. To receive an existing cloud vault, start with a new, empty local vault and approve the download preview. Review both versions after a conflict, and approve each deletion separately. Android sync supports UTF-8 Markdown updates. Incoming changes to existing binary attachments stop for review and preserve the original. Windows also provides the advanced first-connection options.
- **External folder sync (Windows).** Choose a folder and review the proposed scope before allowing synchronization.

## Accounts, network access, and files

- **SmartKeep publishing** requires a SmartKeep account. Sign-in uses account.smartkeep.cn. The publishing service at md.smartkeep.cn receives content and supported linked attachments in the scope you publish and handles access links and feedback.
- **OneDrive sync** requires a Microsoft account. SmartMD uses login.microsoftonline.com for authentication, and Microsoft Graph (graph.microsoft.com) and Microsoft-provided transfer URLs to read and write Markdown and sync metadata in the configured OneDrive app folder. Enable sync only for a vault you intend to connect.
- **AI editing** sends your instructions, relevant conversation history, and selected text or note contents to the model endpoint you configure. If you turn on web search, the configured provider receives a request to perform web search. The provider's terms and pricing apply. Model access can incur charges, and you can use the plugin's other features without configuring an AI provider. AI conversation history and drafts are saved in Obsidian local storage on this device and scoped to the current vault. HTTP sends the request and any configured key without encryption. Use HTTPS for services outside your device.
- **Windows folder sync** reads and writes Markdown in the external folder you choose so it can synchronize changes between that folder and the Obsidian vault. On Windows, the OneDrive setup also inspects local paths to verify the vault location before syncing.
- A SmartKeep or Microsoft account is needed only for the corresponding feature.

## Build

Use Node.js 20 or later with the checked-in npm lockfile.

    npm ci
    npm run build

The build produces main.js. The release also contains manifest.json and styles.css.

## License

SmartMD is licensed under [MIT](LICENSE). It includes [Marked](licenses/marked.txt), uses the Obsidian API, and is an independent project. Obsidian is a trademark of Dynalist Inc. Microsoft and OneDrive are trademarks of Microsoft Corporation.

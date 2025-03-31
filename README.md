# AI-Powered-Translation-Plugin-for-Word-Processing-Applications

This add-in works with any word processing application and helps translate selected text or the entire document using AI. It supports multiple languages and enhances translations based on user preferences.

**Platforms support:** Google Docs and Microsoft Word Online.

**Translation AI model used:** Gemini 2.0 Flash.

**Language packages from:** [LibreTranslate Languages](https://libretranslate.com/languages)

## 1. Development Overview
### 1.1. Google Docs

Open **Google Docs** and navigate to **Extensions > Apps Script**. In the **Apps Script editor**, write the backend logic in the `Code.gs` file and create a `Sidebar.html` file for the frontend interface. Once the coding process is complete, click **"Run"** and select `onOpen()`, or deploy it using the available deployment options to integrate it as a **Google Docs add-on**.

Go back to your Google Docs file, and you'll see a new **"AI Translator"** menu in the toolbar.

### 1.2. Microsoft Word Online

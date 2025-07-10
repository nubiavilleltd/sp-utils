# SharePoint Utils (SPUtils)

A lightweight, dependency-free JavaScript utility for performing CRUD operations on SharePoint lists using the SharePoint REST API.  
No frameworks or external dependencies required. Built for modern SharePoint pages.

---

## 🚀 Features

- 🔹 Get list items with optional filtering
- 🔹 Create new list items
- 🔹 Update existing items
- 🔹 Delete existing items
- 🔹 Upload files to Document Libraries (with optional folder support)
- 🔹 Create folders within Document Libraries
- 🔹 Automatically retrieves request digest and list metadata
- 🔹 Works with modern `async/await`
- 🔹 Lightweight and framework agnostic (no dependencies)

---

## 🌐 CDN Integration

Use the CDN-hosted script via [jsDelivr](https://www.jsdelivr.com/):

```html
<!-- Production-ready version -->
<script src="https://cdn.jsdelivr.net/gh/nubiavilleltd/sp-utils@v1.0.0/utils.min.js"></script>
```

---

## 🔧 Setup

`SPUtils` automatically detects your SharePoint site URL via `_spPageContextInfo`.  
If you need to manually override it (e.g. for cross-site usage), you can use:

```js
SPUtils.setSiteUrl("https://yourdomain.sharepoint.com/sites/yoursite");
```

---

## 📘 API Reference

Each method returns a `Promise`, so you can use them with `async/await` or `.then()`.

### `SPUtils.setSiteUrl(url: string): void`

Set or override the base SharePoint site URL.

| Param | Type   | Description              |
|-------|--------|--------------------------|
| url   | string | Full SharePoint site URL |

```js
SPUtils.setSiteUrl("https://yourcompany.sharepoint.com/sites/team");
```

---

### `SPUtils.getDigest(): Promise<string>`

Returns the current SharePoint request digest token required for write operations.

```js
const digest = await SPUtils.getDigest();
```

---

### `SPUtils.getListMeta(listName: string): Promise<string>`

Retrieves the entity type metadata (`ListItemEntityTypeFullName`) for a given list.

| Param    | Type   | Description                    |
|----------|--------|--------------------------------|
| listName | string | Display name of the list       |

```js
const metaType = await SPUtils.getListMeta("Projects");
```

---

### `SPUtils.createItem(listName: string, body: object): Promise<object>`

Creates a new item in the specified SharePoint list.

| Param    | Type     | Description                   |
|----------|----------|-------------------------------|
| listName | string   | Name of the SharePoint list   |
| body     | object   | Field data for the new item   |

```js
await SPUtils.createItem("Tasks", {
  Title: "Launch Campaign",
  Status: "Pending"
});
```

---

### `SPUtils.getItems(listName: string, filter?: string, top?: number): Promise<object[]>`

Fetches list items with optional filtering and maximum limit.

| Param    | Type     | Description                                          |
|----------|----------|------------------------------------------------------|
| listName | string   | Name of the SharePoint list                          |
| filter   | string   | Optional OData `$filter` query                       |
| top      | number   | Optional maximum number of items to retrieve (default: 4999) |

```js
const items = await SPUtils.getItems("Tasks", "Status eq 'Open'");
```

---

### `SPUtils.updateItem(listName: string, itemId: number, updates: object): Promise<void>`

Updates a SharePoint list item by ID.

| Param    | Type     | Description                            |
|----------|----------|----------------------------------------|
| listName | string   | Name of the SharePoint list            |
| itemId   | number   | ID of the item to update               |
| updates  | object   | Fields and values to be updated        |

```js
await SPUtils.updateItem("Tasks", 3, {
  Status: "Completed"
});
```

---

### `SPUtils.deleteItem(listName: string, itemId: number): Promise<void>`

Deletes an item from a SharePoint list by ID.

| Param    | Type   | Description                  |
|----------|--------|------------------------------|
| listName | string | Name of the SharePoint list  |
| itemId   | number | ID of the item to delete     |

```js
await SPUtils.deleteItem("Tasks", 7);
```

---

### `SPUtils.uploadFileToLibrary(libraryName: string, file: File, folderName?: string): Promise<void>`

Uploads a file to a SharePoint Document Library.

You can specify an optional folder name (inside the document library) where the file should be uploaded. If not provided, the file will be uploaded to the root of the library.

| Param       | Type             | Description                                                  |
|-------------|------------------|--------------------------------------------------------------|
| libraryName | string           | Name of the SharePoint Document Library (e.g., "Documents")  |
| file        | File             | File object to be uploaded                                   |
| folderName  | string (optional)| Name of the folder within the library (e.g., "Resumes")      |

```js
await SPUtils.uploadFileToLibrary("Documents", {},  "Resumes");
```

---


### `SPUtils.createFolder(libraryName: string, folderName: string): Promise<void>`

Create a folder in Sharepoint's Document library

| Param       | Type   | Description                                                  |
|-------------|--------|--------------------------------------------------------------|
| libraryName | string | Name of the SharePoint Document Library (e.g., "Documents")  |
| folderName  | string | Name of the folder to create (e.g., "Applications")          |

```js
await SPUtils.createFolder("Documents", "Applications");
```

---

## 📁 Project Structure

```
sp-utils/
│
├── utils.js         # Development version
├── utils.min.js     # Production minified version
└── README.md        # This documentation
```

---

## ⚠️ Notes

- Works on both classic and modern SharePoint pages.
- Requires appropriate list permissions (Contribute or higher) to perform create, update, or delete operations.
- If you're embedding this in a SharePoint Framework (SPFx) web part, make sure to manually set the site URL via `setSiteUrl()`.

---

## 📜 License

MIT License © [Nubiaville Ltd](https://github.com/nubiavilleltd)
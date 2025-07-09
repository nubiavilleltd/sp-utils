(function (global) {
  const SharePointUtils = {
    siteUrl: (window._spPageContextInfo && _spPageContextInfo.webAbsoluteUrl) || "",

    setSiteUrl(url) {
      this.siteUrl = url || (window._spPageContextInfo && _spPageContextInfo.webAbsoluteUrl);
    },

    async getDigest() {
      return document.getElementById("__REQUESTDIGEST")?.value;
    },

    async getListMeta(listName) {
      const res = await fetch(
        `${this.siteUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(listName)}')?$select=ListItemEntityTypeFullName`,
        { headers: { Accept: "application/json;odata=verbose" } }
      );
      const json = await res.json();
      return json.d.ListItemEntityTypeFullName;
    },

    async createItem(listName, body) {
      const listMeta = await this.getListMeta(listName);
      const digest = await this.getDigest();

      const res = await fetch(
        `${this.siteUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(listName)}')/items`,
        {
          method: "POST",
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": digest,
          },
          body: JSON.stringify({
            "__metadata": { type: listMeta },
            ...body,
          }),
        }
      );

      if (!res.ok) throw new Error(`Create failed: ${await res.text()}`);
      return res.json();
    },

    async getItems(listName, filter = "", top = 4999) {
      let url = `${this.siteUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(listName)}')/items?$top=${top}`;
      if (filter) {
        url += `&$filter=${encodeURIComponent(filter)}`;
      }

      const res = await fetch(url, {
        headers: { Accept: "application/json;odata=verbose" },
      });

      if (!res.ok) throw new Error(`Get failed: ${await res.text()}`);
      return (await res.json()).d.results;
    },

    async updateItem(listName, itemId, updates) {
      const listMeta = await this.getListMeta(listName);
      const digest = await this.getDigest();

      const res = await fetch(
        `${this.siteUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(listName)}')/items(${itemId})`,
        {
          method: "POST",
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": digest,
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*",
          },
          body: JSON.stringify({
            "__metadata": { type: listMeta },
            ...updates,
          }),
        }
      );

      if (!res.ok) throw new Error(`Update failed: ${await res.text()}`);
    },

    async deleteItem(listName, itemId) {
      const digest = await this.getDigest();

      const res = await fetch(
        `${this.siteUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(listName)}')/items(${itemId})`,
        {
          method: "POST",
          headers: {
            Accept: "application/json;odata=verbose",
            "X-RequestDigest": digest,
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE",
          },
        }
      );

      if (!res.ok) throw new Error(`Delete failed: ${await res.text()}`);
    }
  };

  global.SPUtils = SharePointUtils;
})(window);

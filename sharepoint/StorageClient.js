const siteId = "root";
const listId = "todolist";

// Use Microsoft Graph client to manipulate SharePoint list
export class StorageClient {

    graphClient;

    constructor(graphClient) {
        this.graphClient = graphClient
    }

    async getItems() {
        return (await this.graphClient.api(`/sites/${siteId}/lists/${listId}/items?expand=fields`).get()).value.map(item => item.fields);
    }

    async addItem(content) {
        await this.graphClient.api(`/sites/${siteId}/lists/${listId}/items`).post({
            fields: {
                Title: content
            }
        });
    }

    async updateItemContent(id, content) {
        await this.graphClient.api(`/sites/${siteId}/lists/${listId}/items/${id}/fields`).patch({
            Title: content
        });
    }

    async deleteItem(id) {
        await this.graphClient.api(`/sites/${siteId}/lists/${listId}/items/${id}`).delete();
    }

    async updateItemCompltionStatus(id, isComplete) {
        await this.graphClient.api(`/sites/${siteId}/lists/${listId}/items/${id}/fields`).patch({
            isComplete: isComplete
        });
    }
}
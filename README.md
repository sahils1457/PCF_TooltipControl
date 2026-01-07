import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { DynamicDetailsList, IDynamicDetailsListProps } from "./DynamicsDetailsList";
import { GetSampleData } from "./GetSampleData";
import * as React from "react";
import * as ReactDOM from "react-dom";

export class FetchXmlDetailsList implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _detailsList: DynamicDetailsList | null = null;
    private _isTestHarness: boolean = false;
    private _currentPage: number = 1;
    private _pageSize: number = 50;
    private _totalRecords: number = 0;
    
    constructor() {
        this._container = document.createElement("div");
        this._container.style.width = "100%";
        this._container.style.height = "100%";
    }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;
        
        // Check if running in test harness
        this._isTestHarness = this.isTestHarness();
        
        // Access parameters using 'any' to bypass TypeScript restrictions
        const params = context.parameters as any;
        
        // Set page size from input parameter or default
        const itemsPerPage = params.ItemsPerPage?.raw;
        this._pageSize = itemsPerPage ? itemsPerPage : 50;
        
        console.log("FetchXmlDetailsList initialized", {
            isTestHarness: this._isTestHarness,
            pageSize: this._pageSize
        });
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        
        try {
            // Re-check if test harness on each update
            this._isTestHarness = this.isTestHarness();
            
            // Access parameters using 'any' to bypass TypeScript restrictions
            const params = context.parameters as any;
            
            // Check if we have valid FetchXml and ColumnLayout
            const fetchXml = params.FetchXml?.raw || "";
            const columnLayout = params.ColumnLayoutJson?.raw || "";
            
            // If in test harness OR missing required params, use sample data
            if (this._isTestHarness || !fetchXml || !columnLayout || 
                fetchXml.trim() === "" || columnLayout.trim() === "" || 
                columnLayout === "val") {
                console.log("Using sample data (test harness or missing parameters)");
                this.loadSampleData();
            } else {
                this.loadDynamicsData();
            }
        } catch (error) {
            console.error("Error updating view:", error);
            this.showError("Failed to load data: " + (error as Error).message);
        }
    }

    private isTestHarness(): boolean {
        // Check if we're in the test harness by checking for WebApi
        try {
            const webAPI = (this._context as any).webAPI;
            return !webAPI || typeof webAPI.retrieveMultipleRecords !== 'function';
        } catch {
            return true;
        }
    }

    private loadSampleData(): void {
        console.log("Loading sample data for test harness");
        
        const sampleData = GetSampleData();
        const startIndex = (this._currentPage - 1) * this._pageSize;
        const endIndex = startIndex + this._pageSize;
        const paginatedItems = sampleData.dataItems.slice(startIndex, endIndex);
        
        this._totalRecords = sampleData.dataItems.length;
        
        const props: any = {
            context: this._context,
            items: paginatedItems,
            columns: sampleData.columns,
            primaryEntityName: sampleData.primaryEntityName,
            isDebugMode: this.getDebugMode(),
            baseD365Url: this.getBaseEnvironmentUrl(),
            pagination: {
                currentPage: this._currentPage,
                pageSize: this._pageSize,
                totalRecords: this._totalRecords,
                onPageChange: (page: number) => this.handlePageChange(page)
            }
        };
        
        this.renderDetailsList(props);
    }

    private async loadDynamicsData(): Promise<void> {
        const debugMode = this.getDebugMode();
        
        if (debugMode) {
            debugger;
        }

        // Access parameters using 'any' to bypass TypeScript restrictions
        const params = this._context.parameters as any;
        
        // Get parameters
        let fetchXml = params.FetchXml?.raw || "";
        const columnLayoutJson = params.ColumnLayoutJson?.raw || "[]";
        const recordIdPlaceholder = params.RecordIdPlaceholder?.raw || "[RECORDID]";
        const overriddenRecordIdFieldName = params.OverriddenRecordIdFieldName?.raw;

        if (!fetchXml || fetchXml.trim() === "") {
            console.warn("No FetchXml provided, loading sample data instead");
            this.loadSampleData();
            return;
        }

        // Get record ID
        let recordId = this.getRecordId(overriddenRecordIdFieldName || undefined);
        
        if (!recordId) {
            this.showError("Could not determine record ID");
            return;
        }

        // Replace placeholder with actual record ID
        fetchXml = fetchXml.replace(new RegExp(recordIdPlaceholder, 'g'), recordId);

        // Add pagination to FetchXml
        const paginatedFetchXml = this.addPaginationToFetchXml(fetchXml, this._currentPage, this._pageSize);

        if (debugMode) {
            console.log("DynamicDetailsList fetchXml:", paginatedFetchXml);
            console.log("DynamicDetailsList columnLayout:", columnLayoutJson);
        }

        // Parse columns
        let columns: any[];
        try {
            // Clean up the JSON string in case it has issues
            const cleanedJson = columnLayoutJson.trim();
            if (!cleanedJson || cleanedJson === "" || cleanedJson === "val") {
                console.warn("Invalid or empty ColumnLayoutJson, loading sample data");
                this.loadSampleData();
                return;
            }
            columns = JSON.parse(cleanedJson);
        } catch (error) {
            console.error("Failed to parse ColumnLayoutJson:", error);
            console.warn("Loading sample data instead");
            this.loadSampleData();
            return;
        }

        // Execute FetchXml
        try {
            const entityName = this.extractEntityNameFromFetchXml(fetchXml);
            const webAPI = (this._context as any).webAPI;
            const encodedFetchXml = encodeURIComponent(paginatedFetchXml);
            const result = await webAPI.retrieveMultipleRecords(
                entityName,
                `?fetchXml=${encodedFetchXml}`
            );

            if (debugMode) {
                console.log("webAPI.retrieveMultipleRecords result:", result);
            }

            // Get total record count for pagination
            this._totalRecords = await this.getTotalRecordCount(entityName, fetchXml);

            if (debugMode) {
                console.log("Total records:", this._totalRecords);
                console.log("Current page:", this._currentPage);
                console.log("Page size:", this._pageSize);
            }

            const props: any = {
                context: this._context,
                items: result.entities || [],
                columns: columns,
                primaryEntityName: entityName,
                fetchXml: paginatedFetchXml,
                isDebugMode: debugMode,
                baseD365Url: this.getBaseEnvironmentUrl(),
                pagination: {
                    currentPage: this._currentPage,
                    pageSize: this._pageSize,
                    totalRecords: this._totalRecords,
                    onPageChange: (page: number) => this.handlePageChange(page)
                }
            };

            this.renderDetailsList(props);
        } catch (error) {
            console.error("Error executing FetchXml:", error);
            this.showError("Failed to execute FetchXml: " + (error as Error).message);
        }
    }

    private getDebugMode(): boolean {
        const params = this._context.parameters as any;
        const debugValue = params.DebugMode?.raw;
        // DebugMode is an enum: "0" = Off, "1" = On
        return debugValue === "1";
    }

    private getRecordId(overriddenFieldName?: string): string | null {
        try {
            if (overriddenFieldName) {
                // Try to get ID from specified lookup field
                const lookupValue = (this._context as any).parameters[overriddenFieldName];
                if (lookupValue && lookupValue.raw && lookupValue.raw[0]) {
                    return lookupValue.raw[0].id;
                }
            }

            // Try to get current record ID from page context
            const pageContext = (this._context as any).page;
            if (pageContext && pageContext.entityId) {
                return pageContext.entityId;
            }

            // Fallback: try to get from mode
            const mode = this._context.mode as any;
            if (mode && mode.contextInfo && mode.contextInfo.entityId) {
                return mode.contextInfo.entityId;
            }

            return null;
        } catch (error) {
            console.error("Error getting record ID:", error);
            return null;
        }
    }

    private getBaseEnvironmentUrl(): string {
        try {
            // Try multiple methods to get the base URL
            const pageContext = (this._context as any).page;
            
            // Method 1: Try getClientUrl (works in Dynamics)
            if (pageContext && typeof pageContext.getClientUrl === 'function') {
                try {
                    return pageContext.getClientUrl();
                } catch (e) {
                    console.debug("getClientUrl failed, trying alternative methods");
                }
            }
            
            // Method 2: Try to get from context
            if (this._context && (this._context as any).client) {
                const client = (this._context as any).client;
                if (client.getClient && typeof client.getClient === 'function') {
                    const clientInfo = client.getClient();
                    if (clientInfo && clientInfo.baseUrl) {
                        return clientInfo.baseUrl;
                    }
                }
            }
            
            // Method 3: Fallback to window location (works in test harness)
            if (typeof window !== 'undefined' && window.location) {
                return window.location.origin;
            }
            
            // Method 4: Return empty string as last resort
            return "";
        } catch (error) {
            // Silently handle error in test harness
            console.debug("Could not get base environment URL (test harness mode):", error);
            return typeof window !== 'undefined' && window.location ? window.location.origin : "";
        }
    }

    private extractEntityNameFromFetchXml(fetchXml: string): string {
        const match = fetchXml.match(/<entity\s+name=['"]([^'"]+)['"]/);
        return match ? match[1] : "entity";
    }

    private addPaginationToFetchXml(fetchXml: string, page: number, pageSize: number): string {
        // Remove existing count/page attributes
        let xml = fetchXml.replace(/\s+count=['"][^'"]*['"]/g, '');
        xml = xml.replace(/\s+page=['"][^'"]*['"]/g, '');
        xml = xml.replace(/\s+paging-cookie=['"][^'"]*['"]/g, '');

        // Add new pagination attributes to fetch element
        xml = xml.replace(
            /<fetch/,
            `<fetch count="${pageSize}" page="${page}"`
        );

        return xml;
    }

    private async getTotalRecordCount(entityName: string, fetchXml: string): Promise<number> {
        try {
            // Create count query
            let countFetchXml = fetchXml.replace(/<fetch[^>]*>/, '<fetch aggregate="true">');
            countFetchXml = countFetchXml.replace(
                /<entity[^>]*>/,
                (match) => match + `<attribute name="${entityName}id" alias="count" aggregate="count" />`
            );

            const webAPI = (this._context as any).webAPI;
            const encodedFetchXml = encodeURIComponent(countFetchXml);
            const result = await webAPI.retrieveMultipleRecords(
                entityName,
                `?fetchXml=${encodedFetchXml}`
            );
            
            if (result && result.entities && result.entities.length > 0) {
                return parseInt(result.entities[0].count) || 0;
            }
            
            return 0;
        } catch (error) {
            console.warn("Could not get total count:", error);
            return 0;
        }
    }

    private handlePageChange(page: number): void {
        this._currentPage = page;
        this.updateView(this._context);
    }

    private renderDetailsList(props: any): void {
        // Clear container
        while (this._container.firstChild) {
            this._container.removeChild(this._container.firstChild);
        }

        // Create element to render React component
        const element = React.createElement(DynamicDetailsList, props);
        ReactDOM.render(element, this._container);
    }

    private showError(message: string): void {
        this._container.innerHTML = `
            <div style="padding: 20px; background-color: #fef0f0; border: 1px solid #f5c6cb; border-radius: 4px;">
                <strong style="color: #a94442;">Error:</strong>
                <p style="color: #a94442; margin: 10px 0 0 0;">${message}</p>
            </div>
        `;
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        if (this._container) {
            ReactDOM.unmountComponentAtNode(this._container);
        }
    }
}

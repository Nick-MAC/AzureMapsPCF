import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { AzMapPCF, IAzMapPCFProps, IMapPoint } from "./AzMapPCF";
import * as React from "react";

type DataSet = ComponentFramework.PropertyTypes.DataSet;
type EntityRecord = ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;

export class mapPCF implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void = (): void => { return; };
    private selectedRecordId: string | undefined;
    private pendingAction: string | undefined;
    private pendingLatitude: number | undefined;
    private pendingLongitude: number | undefined;
    private pendingRecordId: string | undefined;
    private actionToken = 0;
    private isPagingRequestInFlight = false;
    private lastPageFingerprint = "";
    private hasConfiguredPageSize = false;
    private accumulatedRecordIds: string[] = [];
    private accumulatedRecords: Record<string, EntityRecord> = {};

    private readonly recordIdPropertySetName = "recordIdColumn";
    private readonly latitudePropertySetName = "latitudeColumn";
    private readonly longitudePropertySetName = "longitudeColumn";
    private readonly titlePropertySetName = "titleColumn";

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        context.mode.trackContainerResize(true);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        const dataSet = context.parameters?.sampleDataSet;

        const pageFingerprint = MapPcfHelpers.getPageFingerprint(dataSet);
        if (this.isPagingRequestInFlight && pageFingerprint !== this.lastPageFingerprint) {
            this.isPagingRequestInFlight = false;
        }
        this.lastPageFingerprint = pageFingerprint;

        // The DataSet API is paged; accumulate each page so the control behaves consistently
        // across hosts where loadNextPage either appends or replaces the current page.
        this.syncAccumulatedRecords(dataSet);
        this.ensureAllPagesLoaded(dataSet);

        const allocatedWidth = MapPcfHelpers.normalizeDimension(context.mode.allocatedWidth as number | string | undefined);
        const allocatedHeight = MapPcfHelpers.normalizeDimension(context.mode.allocatedHeight as number | string | undefined);
        const points = MapPcfHelpers.mapRecordsToPoints(this.accumulatedRecords, this.accumulatedRecordIds, {
            recordId: this.recordIdPropertySetName,
            latitude: this.latitudePropertySetName,
            longitude: this.longitudePropertySetName,
            title: this.titlePropertySetName
        });

        const mapStyle = MapPcfHelpers.toNonEmptyTrimmed(context.parameters?.mapStyle?.raw);
        const activeRecordId = MapPcfHelpers.toNonEmptyTrimmed(context.parameters?.activeRecordId?.raw);
        const subscriptionKey = MapPcfHelpers.sanitizeSubscriptionKey(context.parameters?.subscriptionKey?.raw);
        const azureMapsAuthFunctionUrl = MapPcfHelpers.sanitizeUrl(context.parameters?.azureMapsAuthFunctionUrl?.raw);
        const mapDomain = MapPcfHelpers.sanitizeMapDomain(context.parameters?.mapDomain?.raw);
        const authConfigurationError = MapPcfHelpers.getAuthConfigurationError(subscriptionKey, azureMapsAuthFunctionUrl);

        // Feature-visibility flags. TwoOptions inputs default to false when unset, so
        // existing features use opt-out ("hide*") semantics to stay backward compatible,
        // while the newer open-record affordance is opt-in ("show*").
        const hideAddPoint = context.parameters?.hideAddPoint?.raw === true;
        const hideDeletePoint = context.parameters?.hideDeletePoint?.raw === true;
        const hideSearchBar = context.parameters?.hideSearchBar?.raw === true;
        const hideStyleControl = context.parameters?.hideStyleControl?.raw === true;
        const hideClusterControl = context.parameters?.hideClusterControl?.raw === true;
        const hideZoomControls = context.parameters?.hideZoomControls?.raw === true;
        const showOpenRecord = context.parameters?.showOpenRecord?.raw === true;
        const openButtonLabel = MapPcfHelpers.toNonEmptyTrimmed(context.parameters?.openButtonLabel?.raw);

        const props: IAzMapPCFProps = {
            subscriptionKey,
            azureMapsAuthFunctionUrl,
            mapDomain,
            allocatedWidth,
            allocatedHeight,
            mapStyle,
            activeRecordId,
            authConfigurationError,
            points,
            hideAddPoint,
            hideDeletePoint,
            hideSearchBar,
            hideStyleControl,
            hideClusterControl,
            hideZoomControls,
            showOpenRecord,
            openButtonLabel,
            onPointSelected: this.onPointSelected,
            onAddPoint: this.onAddPoint,
            onDeletePoint: this.onDeletePoint,
            onOpenRecord: this.onOpenRecord
        };

        return React.createElement(
            AzMapPCF,
            props
        );
    }

    public getOutputs(): IOutputs {
        // Always return concrete (non-undefined) values. Canvas treats undefined
        // outputs as "blank" and its change-detection often fails to raise OnChange
        // on transitions into/out of blank, which suppresses the event entirely.
        return {
            selectedRecordId: this.selectedRecordId ?? "",
            pendingAction: this.pendingAction ?? "",
            pendingLatitude: this.pendingLatitude ?? 0,
            pendingLongitude: this.pendingLongitude ?? 0,
            pendingRecordId: this.pendingRecordId ?? "",
            actionToken: String(this.actionToken)
        };
    }

    public destroy(): void {
        return;
    }

    private onPointSelected = (recordId: string): void => {
        this.selectedRecordId = recordId;
        this.notifyOutputChanged();
    };

    // Emit an "add" request. Persistence is handled canvas-side via OnChange (Patch).
    private onAddPoint = (latitude: number, longitude: number): void => {
        if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) {
            return;
        }

        this.pendingAction = "add";
        this.pendingLatitude = latitude;
        this.pendingLongitude = longitude;
        this.pendingRecordId = undefined;
        this.actionToken += 1;
        this.notifyOutputChanged();
    };

    // Emit a "delete" request. Persistence is handled canvas-side via OnChange (Remove).
    private onDeletePoint = (recordId: string): void => {
        const trimmedRecordId = MapPcfHelpers.toNonEmptyTrimmed(recordId);
        if (!trimmedRecordId) {
            return;
        }

        this.pendingAction = "delete";
        this.pendingRecordId = trimmedRecordId;
        this.pendingLatitude = undefined;
        this.pendingLongitude = undefined;
        this.actionToken += 1;
        this.notifyOutputChanged();
    };

    // Emit an "open" request. The canvas app reads pendingAction/pendingRecordId in
    // OnChange to run its own navigation Power Fx (the OnSelect-equivalent for a code component).
    private onOpenRecord = (recordId: string): void => {
        const trimmedRecordId = MapPcfHelpers.toNonEmptyTrimmed(recordId);
        if (!trimmedRecordId) {
            return;
        }

        this.pendingAction = "open";
        this.pendingRecordId = trimmedRecordId;
        this.pendingLatitude = undefined;
        this.pendingLongitude = undefined;
        this.actionToken += 1;
        this.notifyOutputChanged();
    };

    private ensureAllPagesLoaded(dataSet: DataSet | undefined): void {
        if (!dataSet?.paging) {
            return;
        }

        if (MapPcfHelpers.isDataSetLoading(dataSet)) {
            return;
        }

        const paging = dataSet.paging as {
            hasNextPage?: boolean;
            loadNextPage?: () => void;
            setPageSize?: (size: number) => void;
        };

        if (!this.hasConfiguredPageSize && typeof paging.setPageSize === "function") {
            try {
                paging.setPageSize(10000);
            } catch {
                // Ignore and continue with provider defaults.
            }
            this.hasConfiguredPageSize = true;
        }

        if (!paging.hasNextPage) {
            return;
        }

        if (this.isPagingRequestInFlight || typeof paging.loadNextPage !== "function") {
            return;
        }

        this.isPagingRequestInFlight = true;
        try {
            paging.loadNextPage();
        } catch {
            this.isPagingRequestInFlight = false;
        }
    }

    private syncAccumulatedRecords(dataSet: DataSet | undefined): void {
        if (!dataSet?.records) {
            this.accumulatedRecordIds = [];
            this.accumulatedRecords = {};
            return;
        }

        const recordIds = MapPcfHelpers.getRecordIds(dataSet);
        const isFirstPage = MapPcfHelpers.isFirstPage(dataSet);
        const shouldReplaceAccumulatedRecords = !dataSet.paging || (isFirstPage && !this.isPagingRequestInFlight);

        if (shouldReplaceAccumulatedRecords) {
            this.accumulatedRecordIds = [];
            this.accumulatedRecords = {};
        }

        for (const recordId of recordIds) {
            if (!this.accumulatedRecordIds.includes(recordId)) {
                this.accumulatedRecordIds.push(recordId);
            }

            const record = dataSet.records[recordId];
            if (record) {
                this.accumulatedRecords[recordId] = record;
            }
        }
    }
}

class MapPcfHelpers {
    public static getPageFingerprint(dataSet: DataSet | undefined): string {
        if (!dataSet) {
            return "";
        }

        const recordIds = this.getRecordIds(dataSet);
        return recordIds.join("|");
    }

    public static getRecordIds(dataSet: DataSet): string[] {
        const fallbackRecordIds = Object.keys(dataSet.records);
        return (dataSet.sortedRecordIds && dataSet.sortedRecordIds.length > 0)
            ? dataSet.sortedRecordIds
            : fallbackRecordIds;
    }

    public static isFirstPage(dataSet: DataSet | undefined): boolean {
        const paging = dataSet?.paging as { hasPreviousPage?: boolean } | undefined;
        return paging?.hasPreviousPage !== true;
    }

    public static isDataSetLoading(dataSet: DataSet | undefined): boolean {
        const loadingState = (dataSet as { loading?: boolean } | undefined)?.loading;
        return loadingState === true;
    }

    public static toNonEmptyTrimmed(value: string | null | undefined): string | undefined {
        if (!value) {
            return undefined;
        }

        const trimmedValue = value.trim();
        return trimmedValue.length > 0 ? trimmedValue : undefined;
    }

    public static normalizeDimension(value: number | string | undefined): number | undefined {
        if (value === undefined || value === null || value === "") {
            return undefined;
        }

        const parsedValue = typeof value === "number" ? value : Number(value);
        if (!Number.isFinite(parsedValue) || parsedValue <= 1) {
            return undefined;
        }

        return parsedValue;
    }

    public static sanitizeSubscriptionKey(value: string | null | undefined): string | undefined {
        const trimmedValue = this.toNonEmptyTrimmed(value);
        if (!trimmedValue) {
            return undefined;
        }

        // Normalize common secret formatting artifacts (extra whitespace/newlines/quotes).
        let cleanedValue = trimmedValue;
        cleanedValue = cleanedValue.replace(/[\r\n\t]/g, "").trim();

        if (cleanedValue.length >= 2) {
            const startsWithDoubleQuote = cleanedValue.startsWith('"');
            const endsWithDoubleQuote = cleanedValue.endsWith('"');
            const startsWithSingleQuote = cleanedValue.startsWith("'");
            const endsWithSingleQuote = cleanedValue.endsWith("'");

            if ((startsWithDoubleQuote && endsWithDoubleQuote) || (startsWithSingleQuote && endsWithSingleQuote)) {
                cleanedValue = cleanedValue.slice(1, -1).trim();
            }
        }

        return cleanedValue.length > 0 ? cleanedValue : undefined;
    }

    public static sanitizeMapDomain(value: string | null | undefined): string | undefined {
        const trimmedValue = this.toNonEmptyTrimmed(value);
        if (!trimmedValue) {
            return undefined;
        }

        let cleanedValue = trimmedValue.toLowerCase();
        cleanedValue = cleanedValue.replace(/^https?:\/\//, "");
        cleanedValue = cleanedValue.replace(/\/$/, "");

        if (cleanedValue === "atlas.microsoft.com" || cleanedValue === "atlas.azure.us") {
            return cleanedValue;
        }

        return undefined;
    }

    public static sanitizeUrl(value: string | null | undefined): string | undefined {
        const trimmedValue = this.toNonEmptyTrimmed(value);
        if (!trimmedValue) {
            return undefined;
        }

        try {
            const parsedUrl = new URL(trimmedValue);
            if (parsedUrl.protocol === "http:" || parsedUrl.protocol === "https:") {
                return parsedUrl.toString();
            }
        } catch {
            return undefined;
        }

        return undefined;
    }

    public static getAuthConfigurationError(
        subscriptionKey: string | undefined,
        azureMapsAuthFunctionUrl: string | undefined
    ): string | undefined {
        const hasSubscriptionKey = !!subscriptionKey;
        const hasAuthFunctionUrl = !!azureMapsAuthFunctionUrl;

        if (hasSubscriptionKey && hasAuthFunctionUrl) {
            return "Provide either Azure Maps Subscription Key or Azure Maps Auth Function URL, not both.";
        }

        if (!hasSubscriptionKey && !hasAuthFunctionUrl) {
            return "Provide Azure Maps Subscription Key or Azure Maps Auth Function URL.";
        }

        return undefined;
    }

    public static mapDataSetToPoints(
        dataSet: DataSet | undefined,
        columns: {
            recordId: string;
            latitude: string;
            longitude: string;
            title: string;
        }
    ): IMapPoint[] {
        if (!dataSet?.records) {
            return [];
        }

        return this.mapRecordsToPoints(dataSet.records, this.getRecordIds(dataSet), columns);
    }

    public static mapRecordsToPoints(
        records: Record<string, EntityRecord>,
        recordIds: string[],
        columns: {
            recordId: string;
            latitude: string;
            longitude: string;
            title: string;
        }
    ): IMapPoint[] {
        if (recordIds.length === 0) {
            return [];
        }

        const points: IMapPoint[] = [];

        for (const recordId of recordIds) {
            const record = records[recordId];
            if (!record) {
                continue;
            }

            const latitude = this.getNumberValue(record, columns.latitude);
            const longitude = this.getNumberValue(record, columns.longitude);

            if (latitude === undefined || longitude === undefined) {
                continue;
            }

            const titleValue = this.getTextValue(record, columns.title);
            const title = this.toNonEmptyTrimmed(titleValue) ?? "Location";
            const mappedRecordId = this.getTextValue(record, columns.recordId);
            const outputRecordId = this.toNonEmptyTrimmed(mappedRecordId) ?? recordId;

            points.push({
                id: outputRecordId,
                latitude,
                longitude,
                title
            });
        }

        return points;
    }

    private static toNumber(value: unknown): number | undefined {
        if (typeof value === "number" && Number.isFinite(value)) {
            return value;
        }

        if (typeof value === "string") {
            const parsedValue = Number(value);
            if (Number.isFinite(parsedValue)) {
                return parsedValue;
            }
        }

        return undefined;
    }

    private static safeGetValue(record: EntityRecord, columnName: string): unknown {
        try {
            return record.getValue(columnName);
        } catch {
            return undefined;
        }
    }

    private static safeGetFormattedValue(record: EntityRecord, columnName: string): string {
        try {
            return record.getFormattedValue(columnName) ?? "";
        } catch {
            return "";
        }
    }

    private static getNumberValue(record: EntityRecord, columnName: string): number | undefined {
        return this.toNumber(this.safeGetValue(record, columnName));
    }

    private static getTextValue(record: EntityRecord, columnName: string): string {
        const formattedValue = this.safeGetFormattedValue(record, columnName);
        if (formattedValue.trim().length > 0) {
            return formattedValue;
        }

        const rawValue = this.safeGetValue(record, columnName);
        if (typeof rawValue === "string" && rawValue.trim().length > 0) {
            return rawValue;
        }

        return "";
    }
}
# Canvas App GeoData Search PCF

- [What this PCF Does](#what-this-pcf-does)
- [Why This Control Exists](#why-this-control-exists)
- [Azure Maps Authentication Options](#azure-maps-authentication-options)
  - [SAS URL Azure Architecture (High Level)](#sas-url-azure-architecture-high-level)
- [How To Use](#how-to-use)
  - [Prereqs](#prereqs)
  - [Part 1: Install the solution zip](#part-1-install-the-solution-zip)
  - [Part 2: Choose and set up authentication](#part-2-choose-and-set-up-authentication)
  - [Part 3: Configure the control in your Canvas app](#part-3-configure-the-control-in-your-canvas-app)
  - [Canvas App Dataset Terminology](#canvas-app-dataset-terminology)

## What this PCF Does
![](docs/media/AzureMapPCF.gif)

PCF features include:
- Azure Maps pin rendering from dataset rows
- Point hover popups and click output (`selectedRecordId`)
- In-map [Style picker control](https://learn.microsoft.com/en-us/azure/azure-maps/choose-map-style#add-the-style-picker-control)
- In-map [clustering](https://samples.azuremaps.com/?search=clustering) toggle (on by default)
- In-map force re-focus on dataset button
- In-map [zoom control](https://learn.microsoft.com/en-us/azure/azure-maps/map-add-controls#add-zoom-control)
- Selection sync support via input/output record ids

## Why This Control Exists
Geospatial controls are unavailable in Power Apps US Government (including GCC), per [Microsoft Learn](https://learn.microsoft.com/en-us/power-apps/maker/canvas-apps/geospatial-overview#prerequisites-for-full-support). So a PCF is required to get a map UX in a Power App in GCC/GCC-H/DOD tenants.

There are already some Azure map controls in the [PCF ecosystem](https://pcf.gallery/search?q=azure+maps), but many are designed primarily around model-driven app patterns or don't provide a clean, canvas-first data contract for table-style `Items` data. This control is intentionally built for canvas app usage where makers typically pass a table/collection and map columns explicitly.

## Azure Maps Authentication Options
For an introduction see https://learn.microsoft.com/en-us/azure/azure-maps/azure-maps-authentication. This control supports two [authentication approaches](https://learn.microsoft.com/en-us/javascript/api/azure-maps-control/atlas.authenticationoptions):

1. **Subscription key** (simpler, less secure for production)
2. **[SAS token URL](https://learn.microsoft.com/en-us/azure/azure-maps/azure-maps-authentication#configuring-application-with-sas-token)** (recommended for production)

| Auth Mode | Advantages | Disadvantages |
| --- | --- | --- |
| Subscription key | Fastest setup for dev/test; no extra Azure resources. | Key is exposed to client; broader access scope; harder production rotation/governance. |
| SAS URL (token broker) | No key in client; short-lived tokens; centralized policy and control in backend. | More setup and moving parts (Function, identity, CORS, ops); higher maintenance overhead. |

> [!NOTE]
> You could extend this further by wrapping the Azure Function call in a Power Platform custom connector and then sharing the custom connector with your Power Apps makers. The customer connector should output the SAS Token to pass into the PCF. Using a custom connector would hide the Azure Function endpoint from the maker and add another authentization layer (this pcf currently works with an anonymous Azure Function). To learn more about this architecture see https://learn.microsoft.com/en-us/power-platform/architecture/reference-architectures/custom-connector-canvas

### SAS URL Azure Architecture (High Level)
When using `azureMapsAuthFunctionUrl`, the architecture is:

1. **Canvas App + PCF control** requests a SAS token from your auth endpoint.
2. **Azure Function token broker** authenticates to Azure Resource Manager using managed identity.
3. Function calls [Azure Maps management API `listSas`](https://learn.microsoft.com/en-us/rest/api/maps-management/accounts/list-sas) to mint a short-lived token.
4. Function returns token payload to the PCF control.
5. PCF initializes Azure Maps Web SDK using SAS auth and the returned token.

Typical Azure resources involved:
- Azure Maps account (the map/data resource).
- Azure Function App (the token broker endpoint).
- [Managed identity with RBAC on Azure Maps](https://learn.microsoft.com/en-us/entra/identity/managed-identities-azure-resources/overview) (used to request SAS tokens).

## How To Use
### Prereqs
1. Power Apps Premium licensing is required for this control. It uses external services (`atlas.microsoft.com` / `atlas.azure.us`), so treat it as a premium canvas app component. See https://learn.microsoft.com/en-us/power-apps/developer/component-framework/overview#licensing.
2. PCF/code components must be allowed in the target environment before the control can be used in Canvas apps. In the Power Platform admin center (PPAC), verify the environment setting that enables code components for canvas apps. See https://learn.microsoft.com/en-us/power-apps/developer/component-framework/component-framework-for-canvas-apps#prerequisites

### Part 1: Install the solution zip
1. Open the [Power Apps maker portal](https://make.powerapps.com/).
2. Go to `Solutions`.
3. Select `Import solution`.
4. Upload `AzureMapPCF/bin/Release/AzureMapPCF_#-#-#-#.zip` and complete the import.

### Part 2: Choose and set up authentication
1. [Create an Azure Maps resource](https://learn.microsoft.com/en-us/azure/azure-maps/quick-demo-map-app)
2. Choose one auth mode:

  **Option A: Subscription key**
  - Retrieve a subscription key (Primary or Secondary).
  - Pass that key to `subscriptionKey`.
  - Use for dev/test or low-risk scenarios.

  **Option B: SAS URL (recommended)**
  - Deploy an auth endpoint (for example Azure Function) that returns a SAS token payload.
  - Configure `azureMapsAuthFunctionUrl` with that endpoint.
  - Leave `subscriptionKey` empty when using SAS URL mode.

### Part 3: Configure the control in your Canvas app
The control expects a bound dataset (PCF `DataSet` input) and a small set of additional PCF property values.

Each row in the dataset should represent one location. Rows with missing or non-numeric latitude/longitude are ignored. `recordIdColumn` should be unique per row to avoid ambiguous selection behavior. The dataset should map these columns:

| Property | Required | Datatype | Notes |
| --- | --- | --- | --- |
| `recordIdColumn` | Yes | Text/Guid/Number (string-convertible unique id) | Used as the stable identity for each pin and for selection sync/output (`selectedRecordId`). |
| `latitudeColumn` | Yes | Decimal number | Valid range is `-90` to `90`. |
| `longitudeColumn` | Yes | Decimal number | Valid range is `-180` to `180`. |
| `titleColumn` | No | Text | Label shown in pin UI (tooltip/popup). |

Other Property Values:

| Property | Direction | Required | Datatype | Notes |
| --- | --- | --- | --- | --- |
| `subscriptionKey` | Input | Conditional | Text | Use for subscription-key auth mode. Leave empty when using `azureMapsAuthFunctionUrl`. |
| `azureMapsAuthFunctionUrl` | Input | Conditional | Text | URL of your token broker endpoint that returns Azure Maps SAS token payload. Leave empty when using `subscriptionKey`. |
| `mapDomain` | Input | No | Text | [Azure Maps domain host](https://learn.microsoft.com/en-us/azure/azure-maps/how-to-use-map-control#azure-government-cloud-support). Use `atlas.microsoft.com` (Azure Commercial / GCC) or `atlas.azure.us` (Azure Government / GCCH, DOD). |
| `mapStyle` | Input | No | Text | Default [Azure Maps style](https://learn.microsoft.com/en-us/azure/azure-maps/choose-map-style#add-the-style-picker-control) name, see docs for supported options |
| `activeRecordId` | Input | No | Text | Record id to force-highlight/center from the canvas app side. |
| `selectedRecordId` | Output | No | Text | Record id emitted when a user clicks a map point. |
| `pendingAction` | Output | No | Text | `"add"` or `"delete"` — the edit the user requested in the map. Read this in `OnChange` to decide whether to `Patch` or `Remove`. |
| `pendingLatitude` | Output | No | Decimal | Latitude of the point to add (set when `pendingAction` is `"add"`). |
| `pendingLongitude` | Output | No | Decimal | Longitude of the point to add (set when `pendingAction` is `"add"`). |
| `pendingRecordId` | Output | No | Text | Record id of the point to delete (set when `pendingAction` is `"delete"`). |
| `actionToken` | Output | No | Text | Monotonic token that changes on every add/delete request. Lets `OnChange` fire reliably even for repeated actions. |

Auth configuration rule:
- Provide **exactly one** of `subscriptionKey` or `azureMapsAuthFunctionUrl`.

### Adding and deleting points
The control can request row changes, but (by design for canvas apps) it does **not** write to your data source itself. Instead it emits output properties and the maker wires `Patch()` / `Remove()` in the control's `OnChange`.

In-map UX:
- **Add a point:** click the **Add point** button (top-right) to enter add mode, then click anywhere on the map. The clicked coordinates are emitted as `pendingLatitude` / `pendingLongitude` with `pendingAction = "add"`. Add mode is one-shot (it turns off after one placement).
- **Delete a point:** click a pin to select it, then click the **Delete point** button. The selected id is emitted as `pendingRecordId` with `pendingAction = "delete"`.

> [!NOTE]
> "Creating the relevant columns" happens at design time: the table and its columns (latitude, longitude, title, etc.) must already exist in your data source. The control populates **values** on a new row; it does not create table schema at runtime.

#### Wiring it up with `OnChange`

The control raises the control's single `OnChange` event whenever an add or delete is requested. Two things make this reliable:

1. **Always-defined outputs.** The control reports concrete values (empty string / `0`) rather than blanks, so the canvas change-detection actually fires `OnChange`. (Returning blank/`undefined` outputs is a common reason `OnChange` never fires in canvas apps.)
2. **The `actionToken` guard.** `OnChange` fires for *every* output change — including selecting a pin (which updates `selectedRecordId`). Because `pendingAction` stays set ("sticky") until the next action, re-running the formula on an unrelated `OnChange` would repeat the last add/delete. `actionToken` increments only on a genuine add/delete click, so gate your logic on it.

Initialize a tracker in **App.OnStart**:

```powerapps
Set(varLastToken, "0")
```

Then set the control's **OnChange** (collection example — swap in your collection/column names):

```powerapps
// Only act when actionToken changes (a real add/delete click), never on incidental OnChange.
If(
    Self.actionToken <> varLastToken,
    Set(varLastToken, Self.actionToken);
    Switch(
        Self.pendingAction,

        "add",
        Collect(
            colLocations,
            {
                recordId:  Text(GUID()),
                latitude:  Self.pendingLatitude,
                longitude: Self.pendingLongitude,
                title:     "New location"
            }
        ),

        "delete",
        RemoveIf( colLocations, recordId = Self.pendingRecordId )
    )
)
```

For a Dataverse table instead of an in-memory collection, swap the bodies for `Patch(YourTable, Defaults(YourTable), {...})` and `Remove(YourTable, LookUp(YourTable, <idColumn> = Self.pendingRecordId))`, keeping the same `actionToken` guard.

After the `Collect`/`RemoveIf` (or `Patch`/`Remove`) updates the bound dataset, the new or removed pin is reflected automatically on the next refresh. Make sure the id you store is text so `recordId = Self.pendingRecordId` compares correctly (the control emits ids as strings); wrap a non-text id with `Text(...)`.

> [!NOTE]
> This PCF expects a bound dataset with mapped `recordIdColumn`, `latitudeColumn`, `longitudeColumn`, and optional `titleColumn` (for pin labels/popups). If your source data does not have clean coordinates, you can extend this pattern with keyword-to-coordinate lookup using [Azure Maps Fuzzy Search](https://learn.microsoft.com/en-us/rest/api/maps/search/get-search-fuzzy).

### Canvas App Dataset Terminology
In canvas app terms, think of the input as a **table** (the same conceptual shape used by a control's `Items` property):
- A dataset is a table.
- Each row is a record.
- Each mapped column points to a field in that record.

Example logical shape:

```text
Table(
  { recordId: "A-1001", latitude: 47.6101, longitude: -122.2015, title: "Seattle HQ" },
  { recordId: "A-1002", latitude: 37.7749, longitude: -122.4194, title: "San Francisco" }
)
```

Set each mapping property by typing the source column name as a string value (for example, `"Latitude"`, `"Longitude"`, `"AccountId"`). These mapping properties are not shown in the standard properties pane for the control. Use the property picker/search, find the PCF mapping property name, and enter the column name string there.
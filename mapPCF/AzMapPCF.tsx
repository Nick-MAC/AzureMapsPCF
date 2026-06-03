import * as React from 'react';
import { Label } from '@fluentui/react-components';
import * as atlas from 'azure-maps-control';
import 'azure-maps-control/dist/atlas.min.css';

export interface IMapPoint {
  id: string;
  latitude: number;
  longitude: number;
  title: string;
}

export interface IAzMapPCFProps {
  subscriptionKey?: string;
  azureMapsAuthFunctionUrl?: string;
  authConfigurationError?: string;
  mapDomain?: string;
  allocatedWidth?: number;
  allocatedHeight?: number;
  mapStyle?: string;
  activeRecordId?: string;
  points?: IMapPoint[];
  onPointSelected?: (recordId: string) => void;
  onAddPoint?: (latitude: number, longitude: number) => void;
  onDeletePoint?: (recordId: string) => void;
}

type MapStyleName =
  | 'road'
  | 'grayscale_light'
  | 'grayscale_dark'
  | 'night'
  | 'road_shaded_relief'
  | 'satellite'
  | 'satellite_road_labels'
  | 'blank'
  | 'high_contrast_dark';

interface MapStyleItem {
  value: MapStyleName;
  label: string;
}

interface MapCameraSnapshot {
  center: atlas.data.Position;
  zoom: number;
}

interface MapControlsProps {
  isStyleControlHovered: boolean;
  isStyleMenuOpen: boolean;
  styleLabel: string;
  styleMenuItems: MapStyleItem[];
  selectedStyle: MapStyleName;
  onStyleControlMouseEnter: () => void;
  onStyleControlMouseLeave: () => void;
  onToggleStyleMenu: () => void;
  onStyleSelected: (style: MapStyleName) => void;
  clusteringEnabled: boolean;
  isClusterControlHovered: boolean;
  onToggleClustering: () => void;
  onClusterControlMouseEnter: () => void;
  onClusterControlMouseLeave: () => void;
  onRefocus: () => void;
  onZoomIn: () => void;
  onZoomOut: () => void;
  isAddModeActive: boolean;
  isAddControlHovered: boolean;
  onToggleAddMode: () => void;
  onAddControlMouseEnter: () => void;
  onAddControlMouseLeave: () => void;
  canDelete: boolean;
  isDeleteControlHovered: boolean;
  onDeleteSelected: () => void;
  onDeleteControlMouseEnter: () => void;
  onDeleteControlMouseLeave: () => void;
}

// Value and style normalization helpers are kept outside the component
// so the class can focus on map lifecycle and event orchestration.
class MapValueHelpers {
  public static toNonEmptyTrimmed(value?: string): string | undefined {
    if (!value) {
      return undefined;
    }

    const trimmedValue = value.trim();
    return trimmedValue.length > 0 ? trimmedValue : undefined;
  }

  public static getValidatedStyle(style: string | undefined, styleMenuItems: MapStyleItem[]): MapStyleName {
    const fallbackStyle: MapStyleName = 'road';
    if (!style) {
      return fallbackStyle;
    }

    const normalizedStyle = style.trim().toLowerCase();
    const styleValues = styleMenuItems.map((item) => item.value);
    return styleValues.includes(normalizedStyle as MapStyleName)
      ? (normalizedStyle as MapStyleName)
      : fallbackStyle;
  }

  public static getValidatedMapDomain(domain?: string): 'atlas.microsoft.com' | 'atlas.azure.us' {
    const normalizedDomain = domain?.trim().toLowerCase();
    return normalizedDomain === 'atlas.azure.us' ? 'atlas.azure.us' : 'atlas.microsoft.com';
  }

  public static getStyleLabel(style: MapStyleName, styleMenuItems: MapStyleItem[]): string {
    const match = styleMenuItems.find((item) => item.value === style);
    return match ? match.label : 'Road';
  }
}

// Shape readers centralize the Azure Maps shape/feature compatibility checks.
class MapShapeReaders {
  public static getShapeProperties(shape: unknown): Record<string, unknown> {
    if (shape && typeof shape === 'object') {
      const candidate = shape as { getProperties?: () => unknown; properties?: unknown };
      if (typeof candidate.getProperties === 'function') {
        const properties = candidate.getProperties();
        if (properties && typeof properties === 'object') {
          return properties as Record<string, unknown>;
        }
      }

      if (candidate.properties && typeof candidate.properties === 'object') {
        return candidate.properties as Record<string, unknown>;
      }
    }

    return {};
  }

  public static getShapePosition(shape: unknown): atlas.data.Position | undefined {
    if (shape && typeof shape === 'object') {
      const atlasShape = shape as { getType?: () => string; getCoordinates?: () => unknown };
      if (typeof atlasShape.getType === 'function' && typeof atlasShape.getCoordinates === 'function') {
        if (atlasShape.getType() === 'Point') {
          const coordinates = atlasShape.getCoordinates();
          if (Array.isArray(coordinates) && coordinates.length >= 2) {
            const coordinateValues = coordinates as unknown[];
            const longitude = coordinateValues[0];
            const latitude = coordinateValues[1];
            if (typeof longitude === 'number' && typeof latitude === 'number') {
              return [longitude, latitude];
            }
          }
        }
      }

      const featureLike = shape as { geometry?: { type?: string; coordinates?: unknown } };
      if (featureLike.geometry?.type === 'Point' && Array.isArray(featureLike.geometry.coordinates)) {
        const coordinateValues = featureLike.geometry.coordinates as unknown[];
        if (coordinateValues.length >= 2) {
          const longitude = coordinateValues[0];
          const latitude = coordinateValues[1];
          if (typeof longitude === 'number' && typeof latitude === 'number') {
            return [longitude, latitude];
          }
        }
      }
    }

    return undefined;
  }
}

// Camera and viewport helpers centralize pan/zoom/fit behaviors.
class MapCameraHelpers {
  public static getPointById(points: IMapPoint[], recordId: string | undefined): IMapPoint | undefined {
    if (!recordId) {
      return undefined;
    }

    return points.find((point) => point.id === recordId);
  }

  public static centerOnPoint(
    map: atlas.Map | null,
    point: IMapPoint,
    clusteringEnabled: boolean,
    clusterMaxZoom: number
  ): void {
    if (!map) {
      return;
    }

    const currentZoom = map.getCamera().zoom ?? 5;
    const targetZoom = clusteringEnabled
      ? Math.max(currentZoom, clusterMaxZoom + 1)
      : currentZoom;

    map.setCamera({
      center: [point.longitude, point.latitude],
      zoom: targetZoom,
      type: 'ease'
    });
  }

  public static centerOnSelectedPoint(
    map: atlas.Map | null,
    points: IMapPoint[],
    selectedPointId: string | undefined,
    clusteringEnabled: boolean,
    clusterMaxZoom: number
  ): void {
    const selectedPoint = this.getPointById(points, selectedPointId);
    if (!selectedPoint) {
      return;
    }

    this.centerOnPoint(map, selectedPoint, clusteringEnabled, clusterMaxZoom);
  }

  public static fitCameraToPoints(map: atlas.Map | null, points: IMapPoint[]): void {
    if (!map || points.length === 0) {
      return;
    }

    if (points.length === 1) {
      const onlyPoint = points[0];
      map.setCamera({
        center: [onlyPoint.longitude, onlyPoint.latitude],
        zoom: 12,
        type: 'ease'
      });
      return;
    }

    let minLatitude = Number.POSITIVE_INFINITY;
    let maxLatitude = Number.NEGATIVE_INFINITY;
    let minLongitude = Number.POSITIVE_INFINITY;
    let maxLongitude = Number.NEGATIVE_INFINITY;

    for (const point of points) {
      minLatitude = Math.min(minLatitude, point.latitude);
      maxLatitude = Math.max(maxLatitude, point.latitude);
      minLongitude = Math.min(minLongitude, point.longitude);
      maxLongitude = Math.max(maxLongitude, point.longitude);
    }

    map.setCamera({
      bounds: [minLongitude, minLatitude, maxLongitude, maxLatitude],
      padding: 60,
      maxZoom: 14,
      type: 'ease'
    });
  }

  public static shouldSkipAutoFit(skipAutoFitOnNextRender: boolean, selectedPointId: string | undefined): boolean {
    if (skipAutoFitOnNextRender) {
      return true;
    }

    return !!selectedPointId;
  }

  public static zoomBy(map: atlas.Map | null, delta: number): void {
    if (!map) {
      return;
    }

    const currentZoom = map.getCamera().zoom ?? 5;
    const nextZoom = Math.min(24, Math.max(0, currentZoom + delta));
    map.setCamera({ zoom: nextZoom, type: 'ease' });
  }

  public static captureCameraSnapshot(map: atlas.Map | null): MapCameraSnapshot | undefined {
    if (!map) {
      return undefined;
    }

    const camera = map.getCamera();
    const center = camera.center;
    const zoom = camera.zoom;
    if (
      Array.isArray(center)
      && center.length >= 2
      && typeof center[0] === 'number'
      && typeof center[1] === 'number'
      && typeof zoom === 'number'
    ) {
      return { center: [center[0], center[1]], zoom };
    }

    return undefined;
  }

  public static zoomIntoCluster(map: atlas.Map | null, event: atlas.MapMouseEvent): void {
    if (!map || !event.shapes || event.shapes.length === 0) {
      return;
    }

    const position = MapShapeReaders.getShapePosition(event.shapes[0]);
    if (!position) {
      return;
    }

    const currentZoom = map.getCamera().zoom ?? 5;
    map.setCamera({ center: position, zoom: Math.min(24, currentZoom + 2), type: 'ease' });
  }
}

function MapControls(props: MapControlsProps): React.ReactElement {
  return (
    <div
      style={{
        position: 'absolute',
        top: 8,
        right: 8,
        zIndex: 10,
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'flex-end',
        gap: 6
      }}
    >
      <div
        onMouseEnter={props.onStyleControlMouseEnter}
        onMouseLeave={props.onStyleControlMouseLeave}
        style={{ position: 'relative', display: 'flex', flexDirection: 'column', alignItems: 'flex-end' }}
      >
        <button
          type="button"
          aria-label="Map style menu"
          onClick={props.onToggleStyleMenu}
          style={{
            display: 'inline-flex',
            alignItems: 'center',
            justifyContent: 'center',
            gap: props.isStyleControlHovered ? 6 : 0,
            background: 'rgba(255,255,255,0.95)',
            border: '1px solid #d1d1d1',
            borderRadius: 4,
            padding: props.isStyleControlHovered ? '6px 8px' : 0,
            width: props.isStyleControlHovered ? 'auto' : 30,
            height: 30,
            cursor: 'pointer',
            whiteSpace: 'nowrap'
          }}
        >
          <svg width="14" height="14" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
            <path
              fill="currentColor"
              d="M12 2 2 7l10 5 10-5-10-5Zm0 7.2L5.7 7 12 3.8 18.3 7 12 9.2ZM4 11.4l8 4 8-4v3L12 19l-8-4.6v-3Zm0 5.1 8 4 8-4v3L12 24l-8-4.5v-3Z"
            />
          </svg>
          {props.isStyleControlHovered && (
            <span style={{ fontSize: 12 }}>{props.styleLabel}</span>
          )}
        </button>

        {props.isStyleMenuOpen && (
          <div
            style={{
              position: 'absolute',
              top: 36,
              right: 0,
              display: 'inline-block',
              background: 'rgba(255,255,255,0.98)',
              border: '1px solid #d1d1d1',
              borderRadius: 4,
              boxShadow: '0 4px 10px rgba(0,0,0,0.15)',
              padding: 4,
              zIndex: 11
            }}
          >
            {props.styleMenuItems.map((item) => (
              <button
                key={item.value}
                type="button"
                onClick={() => props.onStyleSelected(item.value)}
                style={{
                  display: 'block',
                  textAlign: 'left',
                  whiteSpace: 'nowrap',
                  background: item.value === props.selectedStyle ? '#eaf3ff' : 'transparent',
                  border: 'none',
                  borderRadius: 3,
                  padding: '7px 8px',
                  cursor: 'pointer',
                  fontSize: 12
                }}
              >
                {item.label}
              </button>
            ))}
          </div>
        )}
      </div>

      <button
        type="button"
        aria-label="Toggle clustering"
        onClick={props.onToggleClustering}
        onMouseEnter={props.onClusterControlMouseEnter}
        onMouseLeave={props.onClusterControlMouseLeave}
        style={{
          display: 'inline-flex',
          alignItems: 'center',
          justifyContent: 'center',
          gap: props.isClusterControlHovered ? 6 : 0,
          background: props.clusteringEnabled ? '#2f6ab3' : 'rgba(255,255,255,0.95)',
          color: props.clusteringEnabled ? '#ffffff' : '#222222',
          border: '1px solid #d1d1d1',
          borderRadius: 4,
          padding: props.isClusterControlHovered ? '6px 8px' : 0,
          width: props.isClusterControlHovered ? 'auto' : 30,
          height: 30,
          cursor: 'pointer',
          fontSize: 12,
          whiteSpace: 'nowrap'
        }}
      >
        <svg width="14" height="14" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
          <circle cx="8" cy="8" r="4" fill="currentColor" />
          <circle cx="16" cy="9" r="4" fill="currentColor" opacity="0.85" />
          <circle cx="12" cy="16" r="4" fill="currentColor" opacity="0.75" />
        </svg>
        {props.isClusterControlHovered && (
          <span>Clusters: {props.clusteringEnabled ? 'On' : 'Off'}</span>
        )}
      </button>

      <button
        type="button"
        aria-label="Add point"
        aria-pressed={props.isAddModeActive}
        onClick={props.onToggleAddMode}
        onMouseEnter={props.onAddControlMouseEnter}
        onMouseLeave={props.onAddControlMouseLeave}
        style={{
          display: 'inline-flex',
          alignItems: 'center',
          justifyContent: 'center',
          gap: props.isAddControlHovered ? 6 : 0,
          background: props.isAddModeActive ? '#2f6ab3' : 'rgba(255,255,255,0.95)',
          color: props.isAddModeActive ? '#ffffff' : '#222222',
          border: '1px solid #d1d1d1',
          borderRadius: 4,
          padding: props.isAddControlHovered ? '6px 8px' : 0,
          width: props.isAddControlHovered ? 'auto' : 30,
          height: 30,
          cursor: 'pointer',
          fontSize: 12,
          whiteSpace: 'nowrap'
        }}
      >
        <svg width="14" height="14" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
          <path
            fill="currentColor"
            d="M12 2c-3.9 0-7 3.1-7 7 0 5.2 7 13 7 13s7-7.8 7-13c0-3.9-3.1-7-7-7Zm1 7v3h3v2h-3v3h-2v-3H8v-2h3V6h2v3Z"
          />
        </svg>
        {props.isAddControlHovered && (
          <span>{props.isAddModeActive ? 'Click map to add' : 'Add point'}</span>
        )}
      </button>

      {props.canDelete && (
        <button
          type="button"
          aria-label="Delete selected point"
          onClick={props.onDeleteSelected}
          onMouseEnter={props.onDeleteControlMouseEnter}
          onMouseLeave={props.onDeleteControlMouseLeave}
          style={{
            display: 'inline-flex',
            alignItems: 'center',
            justifyContent: 'center',
            gap: props.isDeleteControlHovered ? 6 : 0,
            background: props.isDeleteControlHovered ? '#c62828' : 'rgba(255,255,255,0.95)',
            color: props.isDeleteControlHovered ? '#ffffff' : '#c62828',
            border: '1px solid #d1d1d1',
            borderRadius: 4,
            padding: props.isDeleteControlHovered ? '6px 8px' : 0,
            width: props.isDeleteControlHovered ? 'auto' : 30,
            height: 30,
            cursor: 'pointer',
            fontSize: 12,
            whiteSpace: 'nowrap'
          }}
        >
          <svg width="14" height="14" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
            <path
              fill="currentColor"
              d="M9 3h6l1 2h4v2H4V5h4l1-2Zm-3 6h12l-1 11a2 2 0 0 1-2 2H9a2 2 0 0 1-2-2L6 9Zm4 2v8h2v-8h-2Zm4 0v8h2v-8h-2Z"
            />
          </svg>
          {props.isDeleteControlHovered && (
            <span>Delete point</span>
          )}
        </button>
      )}

      <div
        style={{
          display: 'inline-flex',
          flexDirection: 'column',
          background: 'rgba(255,255,255,0.95)',
          border: '1px solid #d1d1d1',
          borderRadius: 4,
          overflow: 'hidden'
        }}
      >
        <button
          type="button"
          aria-label="Refocus to dataset"
          onClick={props.onRefocus}
          style={{
            width: 30,
            height: 30,
            border: 'none',
            background: 'transparent',
            cursor: 'pointer',
            fontSize: 16,
            lineHeight: '16px'
          }}
          title="Refocus to dataset"
        >
          ◎
        </button>
        <div style={{ height: 1, background: '#d1d1d1' }} />
        <button
          type="button"
          aria-label="Zoom in"
          onClick={props.onZoomIn}
          style={{
            width: 30,
            height: 30,
            border: 'none',
            background: 'transparent',
            cursor: 'pointer',
            fontSize: 18,
            lineHeight: '18px'
          }}
        >
          +
        </button>
        <div style={{ height: 1, background: '#d1d1d1' }} />
        <button
          type="button"
          aria-label="Zoom out"
          onClick={props.onZoomOut}
          style={{
            width: 30,
            height: 30,
            border: 'none',
            background: 'transparent',
            cursor: 'pointer',
            fontSize: 20,
            lineHeight: '18px'
          }}
        >
          -
        </button>
      </div>
    </div>
  );
}

export class AzMapPCF extends React.Component<IAzMapPCFProps> {
  private readonly clusterMaxZoom = 14;
  private readonly styleMenuItems: MapStyleItem[] = [
    { value: 'blank', label: 'Blank' },
    { value: 'grayscale_dark', label: 'Grayscale Dark' },
    { value: 'grayscale_light', label: 'Grayscale Light' },
    { value: 'high_contrast_dark', label: 'High Contrast Dark' },
    { value: 'night', label: 'Night' },
    { value: 'road', label: 'Road' },
    { value: 'road_shaded_relief', label: 'Road Shaded Relief' },
    { value: 'satellite', label: 'Satellite' },
    { value: 'satellite_road_labels', label: 'Satellite Road Labels' }
  ];

  private mapContainerRef: React.RefObject<HTMLDivElement> = React.createRef<HTMLDivElement>();
  private map: atlas.Map | null = null;
  private datasource: atlas.source.DataSource | null = null;
  private selectedStyle: MapStyleName;
  private selectedPointId: string | undefined;
  private clusteringEnabled = true;
  private isStyleMenuOpen = false;
  private isStyleControlHovered = false;
  private isClusterControlHovered = false;
  private isAddModeActive = false;
  private isAddControlHovered = false;
  private isDeleteControlHovered = false;
  private skipAutoFitOnNextRender = false;
  private isClusterFlyoutOpen = false;
  private clusterFlyoutPoints: IMapPoint[] = [];
  private clusterFlyoutAnchor: atlas.data.Position | undefined;
  private runtimeErrorMessage: string | undefined;
  private styleMenuCloseTimeoutId: number | undefined;

  public constructor(props: IAzMapPCFProps) {
    super(props);
    this.selectedStyle = MapValueHelpers.getValidatedStyle(props.mapStyle, this.styleMenuItems);
    this.selectedPointId = MapValueHelpers.toNonEmptyTrimmed(props.activeRecordId);
  }

  // React lifecycle methods
  public componentDidMount(): void {
    this.initializeMap();
  }

  public componentDidUpdate(prevProps: IAzMapPCFProps): void {
    this.syncMapConnection(prevProps);
    this.syncStyle(prevProps);
    this.syncSelection(prevProps);
    this.syncResize(prevProps);
    this.syncPoints(prevProps);
  }

  public componentWillUnmount(): void {
    this.clearStyleMenuCloseTimeout();
    this.disposeMap();
  }

  private setRuntimeError(error: unknown, context: string): void {
    const message = error instanceof Error ? error.message : String(error);
    const nextMessage = `Map failed (${context}): ${message}`;
    if (this.runtimeErrorMessage !== nextMessage) {
      this.runtimeErrorMessage = nextMessage;
      this.disposeMap();
      this.forceUpdate();
    }
  }

  private clearRuntimeError(): void {
    if (this.runtimeErrorMessage) {
      this.runtimeErrorMessage = undefined;
      this.forceUpdate();
    }
  }

  // Prop synchronization methods
  private syncMapConnection(prevProps: IAzMapPCFProps): void {
    if (
      prevProps.subscriptionKey !== this.props.subscriptionKey
      || prevProps.azureMapsAuthFunctionUrl !== this.props.azureMapsAuthFunctionUrl
      || prevProps.mapDomain !== this.props.mapDomain
    ) {
      this.disposeMap();
      this.initializeMap();
    }
  }

  private syncStyle(prevProps: IAzMapPCFProps): void {
    if (prevProps.mapStyle !== this.props.mapStyle) {
      this.selectedStyle = MapValueHelpers.getValidatedStyle(this.props.mapStyle, this.styleMenuItems);
      this.applyStyle(this.selectedStyle);
    }
  }

  private syncSelection(prevProps: IAzMapPCFProps): void {
    if (prevProps.activeRecordId !== this.props.activeRecordId) {
      this.selectedPointId = MapValueHelpers.toNonEmptyTrimmed(this.props.activeRecordId);
      this.isClusterFlyoutOpen = false;
      this.clusterFlyoutPoints = [];
      this.clusterFlyoutAnchor = undefined;
      MapCameraHelpers.centerOnSelectedPoint(
        this.map,
        this.props.points ?? [],
        this.selectedPointId,
        this.clusteringEnabled,
        this.clusterMaxZoom
      );
      this.renderPoints();
    }
  }

  private syncResize(prevProps: IAzMapPCFProps): void {
    if (
      this.map &&
      (prevProps.allocatedWidth !== this.props.allocatedWidth || prevProps.allocatedHeight !== this.props.allocatedHeight)
    ) {
      this.map.resize();
    }
  }

  private syncPoints(prevProps: IAzMapPCFProps): void {
    if (prevProps.points !== this.props.points) {
      this.renderPoints();
      if (this.selectedPointId) {
        MapCameraHelpers.centerOnSelectedPoint(
          this.map,
          this.props.points ?? [],
          this.selectedPointId,
          this.clusteringEnabled,
          this.clusterMaxZoom
        );
      } else {
        this.autoFitToDataIfApplicable();
      }
    }
  }

  // Control UI state + interactions
  private applyStyle(style: MapStyleName): void {
    if (!this.map) {
      return;
    }

    this.map.setStyle({ style, language: 'en-US' });
  }

  private toggleStyleMenu = (): void => {
    this.clearStyleMenuCloseTimeout();
    this.isStyleMenuOpen = !this.isStyleMenuOpen;
    this.forceUpdate();
  };

  private clearStyleMenuCloseTimeout(): void {
    if (this.styleMenuCloseTimeoutId !== undefined) {
      window.clearTimeout(this.styleMenuCloseTimeoutId);
      this.styleMenuCloseTimeoutId = undefined;
    }
  }

  private setStyleControlHovered = (hovered: boolean): void => {
    if (hovered) {
      this.clearStyleMenuCloseTimeout();
    }

    if (this.isStyleControlHovered !== hovered) {
      this.isStyleControlHovered = hovered;
      this.forceUpdate();
    }
  };

  private onStyleControlMouseLeave = (): void => {
    this.clearStyleMenuCloseTimeout();
    this.styleMenuCloseTimeoutId = window.setTimeout(() => {
      this.styleMenuCloseTimeoutId = undefined;
      const hadChanges = this.isStyleControlHovered || this.isStyleMenuOpen;
      this.isStyleControlHovered = false;
      this.isStyleMenuOpen = false;
      if (hadChanges) {
        this.forceUpdate();
      }
    }, 200);
  };

  private onStyleSelected = (style: MapStyleName): void => {
    this.clearStyleMenuCloseTimeout();
    const nextStyle = MapValueHelpers.getValidatedStyle(style, this.styleMenuItems);
    this.selectedStyle = nextStyle;
    this.isStyleMenuOpen = false;
    this.applyStyle(nextStyle);
    this.forceUpdate();
  };

  private closeClusterFlyout = (): void => {
    if (!this.isClusterFlyoutOpen && this.clusterFlyoutPoints.length === 0) {
      return;
    }

    this.isClusterFlyoutOpen = false;
    this.clusterFlyoutPoints = [];
    this.clusterFlyoutAnchor = undefined;
    this.forceUpdate();
  };

  private mapClusterLeavesToPoints(leaves: unknown[]): IMapPoint[] {
    const points: IMapPoint[] = [];
    const seenIds = new Set<string>();

    for (const leaf of leaves) {
      const properties = MapShapeReaders.getShapeProperties(leaf);
      const idValue = properties.id;
      const titleValue = properties.title;
      const position = MapShapeReaders.getShapePosition(leaf);

      if (typeof idValue !== 'string' || idValue.trim().length === 0 || !position) {
        continue;
      }

      if (seenIds.has(idValue)) {
        continue;
      }

      seenIds.add(idValue);
      points.push({
        id: idValue,
        title: typeof titleValue === 'string' && titleValue.trim().length > 0 ? titleValue : 'Location',
        longitude: position[0],
        latitude: position[1]
      });
    }

    return points;
  }

  private arePointsAtSameCoordinates(points: IMapPoint[]): boolean {
    if (points.length <= 1) {
      return false;
    }

    const firstPoint = points[0];
    return points.every((point) => point.latitude === firstPoint.latitude && point.longitude === firstPoint.longitude);
  }

  private selectPointById(recordId: string): void {
    this.selectedPointId = recordId;
    this.isClusterFlyoutOpen = false;
    this.clusterFlyoutPoints = [];
    this.clusterFlyoutAnchor = undefined;
    this.renderPoints();
    MapCameraHelpers.centerOnSelectedPoint(
      this.map,
      this.props.points ?? [],
      this.selectedPointId,
      this.clusteringEnabled,
      this.clusterMaxZoom
    );

    if (this.props.onPointSelected) {
      try {
        this.props.onPointSelected(recordId);
      } catch (error) {
        this.setRuntimeError(error, 'point selected callback');
      }
    }

    // Re-render so selection-dependent controls (e.g. the Delete button) update.
    this.forceUpdate();
  }

  private onClusterClicked = (event: atlas.MapMouseEvent): void => {
    void this.handleClusterClickedAsync(event);
  };

  private async handleClusterClickedAsync(event: atlas.MapMouseEvent): Promise<void> {
    if (!this.datasource || !event.shapes || event.shapes.length === 0) {
      return;
    }

    const properties = MapShapeReaders.getShapeProperties(event.shapes[0]);
    const clusterIdValue = properties.cluster_id;
    const pointCountValue = properties.point_count;
    const clickedClusterPosition = MapShapeReaders.getShapePosition(event.shapes[0]);
    const clusterId = typeof clusterIdValue === 'number' ? clusterIdValue : undefined;
    const pointCount = typeof pointCountValue === 'number' ? pointCountValue : 0;

    if (clusterId === undefined || pointCount <= 0) {
      MapCameraHelpers.zoomIntoCluster(this.map, event);
      return;
    }

    try {
      const leaves = await this.datasource.getClusterLeaves(clusterId, Math.min(pointCount, 100), 0);
      const clusterPoints = this.mapClusterLeavesToPoints(leaves as unknown[]);

      if (clusterPoints.length === 0) {
        MapCameraHelpers.zoomIntoCluster(this.map, event);
        return;
      }

      if (!this.arePointsAtSameCoordinates(clusterPoints)) {
        this.closeClusterFlyout();
        MapCameraHelpers.zoomIntoCluster(this.map, event);
        return;
      }

      this.clusterFlyoutPoints = clusterPoints;
      this.clusterFlyoutAnchor = clickedClusterPosition ?? [clusterPoints[0].longitude, clusterPoints[0].latitude];
      this.isClusterFlyoutOpen = true;
      this.forceUpdate();
    } catch (error) {
      this.setRuntimeError(error, 'cluster leaves');
    }
  }

  // Selection and camera methods
  private autoFitToDataIfApplicable(): void {
    if (MapCameraHelpers.shouldSkipAutoFit(this.skipAutoFitOnNextRender, this.selectedPointId)) {
      if (this.skipAutoFitOnNextRender) {
        this.skipAutoFitOnNextRender = false;
      }
      return;
    }

    MapCameraHelpers.fitCameraToPoints(this.map, this.props.points ?? []);
  }

  private onZoomIn = (): void => {
    MapCameraHelpers.zoomBy(this.map, 1);
  };

  private onZoomOut = (): void => {
    MapCameraHelpers.zoomBy(this.map, -1);
  };

  private onRefocus = (): void => {
    this.isClusterFlyoutOpen = false;
    this.clusterFlyoutPoints = [];
    this.clusterFlyoutAnchor = undefined;
    MapCameraHelpers.fitCameraToPoints(this.map, this.props.points ?? []);
    this.forceUpdate();
  };

  private toggleClustering = (): void => {
    const cameraSnapshot = MapCameraHelpers.captureCameraSnapshot(this.map);

    this.clusteringEnabled = !this.clusteringEnabled;
    this.isClusterFlyoutOpen = false;
    this.clusterFlyoutPoints = [];
    this.clusterFlyoutAnchor = undefined;
    this.disposeMap();
    this.initializeMap(cameraSnapshot);
    this.forceUpdate();
  };

  private setClusterControlHovered = (hovered: boolean): void => {
    if (this.isClusterControlHovered !== hovered) {
      this.isClusterControlHovered = hovered;
      this.forceUpdate();
    }
  };

  private setAddControlHovered = (hovered: boolean): void => {
    if (this.isAddControlHovered !== hovered) {
      this.isAddControlHovered = hovered;
      this.forceUpdate();
    }
  };

  private setDeleteControlHovered = (hovered: boolean): void => {
    if (this.isDeleteControlHovered !== hovered) {
      this.isDeleteControlHovered = hovered;
      this.forceUpdate();
    }
  };

  private setAddMode(active: boolean): void {
    this.isAddModeActive = active;
    if (this.map) {
      this.map.getCanvasContainer().style.cursor = active ? 'crosshair' : '';
    }
  }

  private toggleAddMode = (): void => {
    this.setAddMode(!this.isAddModeActive);
    // Adding and the cluster flyout are mutually exclusive interactions.
    this.isClusterFlyoutOpen = false;
    this.clusterFlyoutPoints = [];
    this.clusterFlyoutAnchor = undefined;
    this.forceUpdate();
  };

  private onMapClicked = (event: atlas.MapMouseEvent): void => {
    if (!this.isAddModeActive) {
      return;
    }

    const position = event.position;
    if (!Array.isArray(position) || position.length < 2) {
      return;
    }

    const longitude = position[0];
    const latitude = position[1];
    if (typeof longitude !== 'number' || typeof latitude !== 'number') {
      return;
    }

    // Add mode is one-shot: place the point and return to normal interaction.
    this.setAddMode(false);
    this.forceUpdate();

    if (this.props.onAddPoint) {
      try {
        this.props.onAddPoint(latitude, longitude);
      } catch (error) {
        this.setRuntimeError(error, 'add point callback');
      }
    }
  };

  private onDeleteSelected = (): void => {
    const recordId = this.selectedPointId;
    if (!recordId) {
      return;
    }

    if (this.props.onDeletePoint) {
      try {
        this.props.onDeletePoint(recordId);
      } catch (error) {
        this.setRuntimeError(error, 'delete point callback');
      }
    }

    this.selectedPointId = undefined;
    this.isDeleteControlHovered = false;
    this.renderPoints();
    this.forceUpdate();
  };

  private zoomIntoCluster = (event: atlas.MapMouseEvent): void => {
    MapCameraHelpers.zoomIntoCluster(this.map, event);
  };

  private async fetchSasTokenFromAuthFunction(url: string): Promise<string> {
    const headers: Record<string, string> = {
      Accept: 'application/json'
    };

    const response = await fetch(url, {
      method: 'GET',
      headers
    });

    if (!response.ok) {
      const responseText = await response.text();
      console.error('[AzMapPCF] Auth function error response:', responseText);
      throw new Error(`Auth function failed (${response.status}): ${responseText}`);
    }

    const contentType = response.headers.get('content-type') ?? '';
    if (contentType.toLowerCase().includes('application/json')) {
      const payload = await response.json() as { token?: string; accountSasToken?: string; sasToken?: string };
      const token = payload.token ?? payload.accountSasToken ?? payload.sasToken;
      if (token && token.trim().length > 0) {
        return token.trim();
      }
    }

    const tokenText = (await response.text()).trim();
    if (tokenText.length === 0) {
      throw new Error('Auth function response did not include a token value.');
    }
    return tokenText;
  }

  // Map creation, event wiring, and data rendering
  private initializeMap(cameraSnapshot?: MapCameraSnapshot): void {
    if (!this.mapContainerRef.current || this.props.authConfigurationError) {
      return;
    }

    const usingSubscriptionKey = !!this.props.subscriptionKey;
    const usingAuthFunction = !!this.props.azureMapsAuthFunctionUrl;
    if (!usingSubscriptionKey && !usingAuthFunction) {
      return;
    }

    const center = cameraSnapshot?.center ?? [-77.0369, 38.9072];
    const zoom = cameraSnapshot?.zoom ?? 5;
    const style = this.selectedStyle;
    const mapDomain = MapValueHelpers.getValidatedMapDomain(this.props.mapDomain);
    this.skipAutoFitOnNextRender = !!cameraSnapshot;

    try {
      atlas.setDomain(mapDomain);

      const authOptions = usingSubscriptionKey
        ? {
          authType: atlas.AuthenticationType.subscriptionKey,
          subscriptionKey: this.props.subscriptionKey
        }
        : {
          authType: atlas.AuthenticationType.sas,
          getToken: (resolve: (token: string) => void, reject: (error: string) => void): void => {
            const authFunctionUrl = this.props.azureMapsAuthFunctionUrl;
            if (!authFunctionUrl) {
              console.error('[AzMapPCF] No auth function URL provided');
              reject('Azure Maps auth function URL is required for jwt-sas mode.');
              return;
            }

            const handleTokenFetch = async (): Promise<void> => {
              try {
                const token = await this.fetchSasTokenFromAuthFunction(authFunctionUrl);
                resolve(token);
              } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                console.error('[AzMapPCF] Token fetch failed:', message);
                reject(message);
              }
            };
            void handleTokenFetch();
          }
        };

      this.map = new atlas.Map(this.mapContainerRef.current, {
        authOptions,
        center,
        zoom,
        style,
        language: 'en-US'
      });
      this.clearRuntimeError();
    } catch (error) {
      console.error('[AzMapPCF] Error during map initialization:', error);
      this.setRuntimeError(error, 'create map');
      return;
    }

    this.map.events.add('ready', () => {
      try {
        if (!this.map) {
          return;
        }

      const dataSourceOptions: atlas.DataSourceOptions = {
        cluster: this.clusteringEnabled,
        clusterRadius: 45,
        clusterMaxZoom: this.clusterMaxZoom
      };

      if (this.clusteringEnabled) {
        // Track whether a cluster contains the currently selected point.
        (dataSourceOptions as atlas.DataSourceOptions & { clusterProperties?: unknown }).clusterProperties = {
          selectedCount: ['+', ['case', ['boolean', ['get', 'isSelected'], false], 1, 0]]
        };
      }

      this.datasource = new atlas.source.DataSource(undefined, dataSourceOptions);
      this.map.sources.add(this.datasource);

      const pointLayer = new atlas.layer.SymbolLayer(this.datasource, undefined, {
        filter: this.clusteringEnabled ? ['!', ['has', 'point_count']] : undefined,
        iconOptions: {
          image: ['case', ['boolean', ['get', 'isSelected'], false], 'pin-round-red', 'pin-round-darkblue'],
          anchor: 'bottom',
          offset: [0, -2],
          allowOverlap: true
        },
        textOptions: {
          textField: ['get', 'title'],
          anchor: 'top'
        }
      });

      if (this.clusteringEnabled) {
        const clusterBubbleLayer = new atlas.layer.BubbleLayer(this.datasource, undefined, {
          filter: ['has', 'point_count'],
          radius: ['step', ['get', 'point_count'], 18, 20, 22, 100, 28, 500, 34],
          color: [
            'case',
            ['>', ['get', 'selectedCount'], 0],
            '#c62828',
            ['step', ['get', 'point_count'], '#4f83cc', 20, '#2f6ab3', 100, '#204d88', 500, '#13335d']
          ],
          strokeColor: '#ffffff',
          strokeWidth: 1
        });

        const clusterCountLayer = new atlas.layer.SymbolLayer(this.datasource, undefined, {
          filter: ['has', 'point_count'],
          textOptions: {
            textField: ['get', 'point_count_abbreviated'],
            color: '#ffffff',
            size: 12
          }
        });

        this.map.layers.add([clusterBubbleLayer, clusterCountLayer, pointLayer]);
        this.map.events.add('click', clusterBubbleLayer, this.onClusterClicked);
        this.map.events.add('click', clusterCountLayer, this.onClusterClicked);
        this.map.events.add('mouseover', clusterBubbleLayer, () => {
          if (!this.map) {
            return;
          }

          this.map.getCanvasContainer().style.cursor = 'pointer';
        });
        this.map.events.add('mouseover', clusterCountLayer, () => {
          if (!this.map) {
            return;
          }

          this.map.getCanvasContainer().style.cursor = 'pointer';
        });
        this.map.events.add('mouseout', clusterBubbleLayer, () => {
          if (!this.map) {
            return;
          }

          this.map.getCanvasContainer().style.cursor = '';
        });
        this.map.events.add('mouseout', clusterCountLayer, () => {
          if (!this.map) {
            return;
          }

          this.map.getCanvasContainer().style.cursor = '';
        });
      } else {
        this.map.layers.add(pointLayer);
      }

      const popup = new atlas.Popup({
        closeButton: false,
        pixelOffset: [0, -18]
      });

      this.map.events.add('mouseover', pointLayer, (event: atlas.MapMouseEvent) => {
        if (!this.map || !event.shapes || event.shapes.length === 0 || this.isAddModeActive) {
          return;
        }

        const shape = event.shapes[0];
        const properties = MapShapeReaders.getShapeProperties(shape);
        const title = typeof properties?.title === 'string' && properties.title.length > 0
          ? properties.title
          : 'Location';
        const geometry = MapShapeReaders.getShapePosition(shape);

        if (!geometry) {
          return;
        }

        popup.setOptions({
          content: `<div style="padding:8px 10px; font-size:12px;">${title}</div>`,
          position: geometry
        });
        popup.open(this.map);

        this.map.getCanvasContainer().style.cursor = 'pointer';
      });

      this.map.events.add('mouseout', pointLayer, () => {
        if (!this.map) {
          return;
        }

        popup.close();
        this.map.getCanvasContainer().style.cursor = this.isAddModeActive ? 'crosshair' : '';
      });

      // Map-level click handles dropping a new point while add mode is active.
      this.map.events.add('click', this.onMapClicked);

      this.map.events.add('click', pointLayer, (event: atlas.MapMouseEvent) => {
        if (!event.shapes || event.shapes.length === 0 || this.isAddModeActive) {
          return;
        }

        const shape = event.shapes[0];
        const properties = MapShapeReaders.getShapeProperties(shape);
        const recordId = properties?.id;
        if (typeof recordId === 'string' && recordId.length > 0) {
          this.selectPointById(recordId);
        }
      });

        this.renderPoints();
        if (this.selectedPointId) {
          MapCameraHelpers.centerOnSelectedPoint(
            this.map,
            this.props.points ?? [],
            this.selectedPointId,
            this.clusteringEnabled,
            this.clusterMaxZoom
          );
        } else {
          this.autoFitToDataIfApplicable();
        }
      } catch (error) {
        this.setRuntimeError(error, 'map ready');
      }
    });
  }

  private renderPoints(): void {
    if (!this.datasource) {
      return;
    }

    this.datasource.clear();

    const points = this.props.points ?? [];
    if (points.length === 0) {
      return;
    }

    const features = points.map((point) => new atlas.data.Feature(
      new atlas.data.Point([point.longitude, point.latitude]),
      {
        title: point.title,
        id: point.id,
        isSelected: this.selectedPointId === point.id
      }
    ));

    this.datasource.add(features);
  }

  private disposeMap(): void {
    if (this.map) {
      this.map.dispose();
      this.map = null;
    }

    this.datasource = null;
  }

  private renderClusterFlyout(): React.ReactNode {
    if (!this.isClusterFlyoutOpen || this.clusterFlyoutPoints.length === 0 || !this.map || !this.clusterFlyoutAnchor) {
      return null;
    }

    const mapContainer = this.mapContainerRef.current;
    const containerWidth = mapContainer?.clientWidth ?? 0;
    const containerHeight = mapContainer?.clientHeight ?? 0;
    const panelWidth = 240;
    const panelMaxHeight = 260;
    const anchorPixels = this.map.positionsToPixels([this.clusterFlyoutAnchor])[0];
    const anchorX = typeof anchorPixels?.[0] === 'number' ? anchorPixels[0] : 0;
    const anchorY = typeof anchorPixels?.[1] === 'number' ? anchorPixels[1] : 0;

    const desiredLeft = anchorX + 12;
    const desiredTop = anchorY - 12;
    const maxLeft = Math.max(8, containerWidth - panelWidth - 8);
    const maxTop = Math.max(8, containerHeight - panelMaxHeight - 8);
    const flyoutLeft = Math.min(Math.max(8, desiredLeft), maxLeft);
    const flyoutTop = Math.min(Math.max(8, desiredTop), maxTop);

    return (
      <div
        style={{
          position: 'absolute',
          top: flyoutTop,
          left: flyoutLeft,
          zIndex: 12,
          width: panelWidth,
          maxHeight: panelMaxHeight,
          overflowY: 'auto',
          background: 'rgba(255,255,255,0.98)',
          border: '1px solid #d1d1d1',
          borderRadius: 6,
          boxShadow: '0 6px 16px rgba(0,0,0,0.2)',
          padding: 8
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 6 }}>
          <span style={{ fontSize: 12, fontWeight: 600 }}>
            Select a document ({this.clusterFlyoutPoints.length})
          </span>
          <button
            type="button"
            aria-label="Close cluster list"
            onClick={this.closeClusterFlyout}
            style={{ border: 'none', background: 'transparent', cursor: 'pointer', fontSize: 16, lineHeight: '16px' }}
          >
            x
          </button>
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
          {this.clusterFlyoutPoints.map((point) => (
            <button
              key={point.id}
              type="button"
              onClick={() => this.selectPointById(point.id)}
              style={{
                textAlign: 'left',
                border: '1px solid #d9d9d9',
                borderRadius: 4,
                background: '#ffffff',
                padding: '6px 8px',
                cursor: 'pointer',
                fontSize: 12
              }}
            >
              <div
                title={point.title}
                style={{
                  fontWeight: 600,
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  whiteSpace: 'nowrap'
                }}
              >
                {point.title}
              </div>
              <div
                title={point.id}
                style={{
                  color: '#666666',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  whiteSpace: 'nowrap'
                }}
              >
                {point.id}
              </div>
            </button>
          ))}
        </div>
      </div>
    );
  }

  // React render output
  public render(): React.ReactNode {
    if (this.runtimeErrorMessage) {
      return (
        <Label>
          {this.runtimeErrorMessage}
        </Label>
      );
    }

    if (this.props.authConfigurationError) {
      return (
        <Label>
          {this.props.authConfigurationError}
        </Label>
      );
    }

    const heightStyle = this.props.allocatedHeight && this.props.allocatedHeight > 1
      ? `${this.props.allocatedHeight}px`
      : '100%';

    return (
      <div style={{ position: 'relative', width: '100%', height: heightStyle }}>
        {this.renderClusterFlyout()}
        <MapControls
          isStyleControlHovered={this.isStyleControlHovered}
          isStyleMenuOpen={this.isStyleMenuOpen}
          styleLabel={MapValueHelpers.getStyleLabel(this.selectedStyle, this.styleMenuItems)}
          styleMenuItems={this.styleMenuItems}
          selectedStyle={this.selectedStyle}
          onStyleControlMouseEnter={() => this.setStyleControlHovered(true)}
          onStyleControlMouseLeave={this.onStyleControlMouseLeave}
          onToggleStyleMenu={this.toggleStyleMenu}
          onStyleSelected={this.onStyleSelected}
          clusteringEnabled={this.clusteringEnabled}
          isClusterControlHovered={this.isClusterControlHovered}
          onToggleClustering={this.toggleClustering}
          onClusterControlMouseEnter={() => this.setClusterControlHovered(true)}
          onClusterControlMouseLeave={() => this.setClusterControlHovered(false)}
          onRefocus={this.onRefocus}
          onZoomIn={this.onZoomIn}
          onZoomOut={this.onZoomOut}
          isAddModeActive={this.isAddModeActive}
          isAddControlHovered={this.isAddControlHovered}
          onToggleAddMode={this.toggleAddMode}
          onAddControlMouseEnter={() => this.setAddControlHovered(true)}
          onAddControlMouseLeave={() => this.setAddControlHovered(false)}
          canDelete={!!this.selectedPointId && !!this.props.onDeletePoint}
          isDeleteControlHovered={this.isDeleteControlHovered}
          onDeleteSelected={this.onDeleteSelected}
          onDeleteControlMouseEnter={() => this.setDeleteControlHovered(true)}
          onDeleteControlMouseLeave={() => this.setDeleteControlHovered(false)}
        />
        <div ref={this.mapContainerRef} style={{ width: '100%', height: '100%' }} />
      </div>
    );
  }
}

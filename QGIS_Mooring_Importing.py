"""Import a Main.py mooring-coordinate CSV into QGIS layers.

The seismic/PyVista screen belongs to Main.py's Unified Dashboard. This script
performs only the QGIS coordinate/GPKG import, so users can continue to inspect
and work with the exported layers in QGIS.
"""

import gc
import os
import time

import processing
from qgis.core import (
    QgsLineSymbol,
    QgsMarkerSymbol,
    QgsProject,
    QgsProperty,
    QgsSymbolLayer,
    QgsVectorFileWriter,
    QgsVectorLayer,
    QgsWkbTypes,
)


try:
  from qgis.utils import iface
  if iface:
    iface.actionShowPythonDialog().trigger()
    print("[INFO] Python console opened automatically.")
except Exception as exc:
  print(f"[WARNING] Could not open Python console automatically: {exc}")


# =====================================================================
# DYNAMIC PROJECT & PATH CONFIGURATION FROM Main.py
# =====================================================================
project_name = os.environ.get("PROJECT_NAME", "Stonehaven")
project_root = os.environ.get(
    "PROJECT_ROOT", r"C:\Users\r02an25\Documents\Floating_Offshore_Wind"
)
print(f"[INFO] Active Project Name: '{project_name}'")
print(f"[INFO] Active Project Root: '{project_root}'")

raw_env_path = os.environ.get("MOORING_CSV_PATH")
print(f"[DEBUG] Raw MOORING_CSV_PATH received from Main.py: '{raw_env_path}'")
if raw_env_path:
  clean_env_path = raw_env_path.strip().strip("'\"")
  if os.path.exists(clean_env_path):
    csv_path = clean_env_path.replace("\\", "/")
    print(f"[INFO] Validated dynamic CSV path: {csv_path}")
  else:
    print(f"[ERROR] Path received, but file not found on disk: {clean_env_path}")
    csv_path = (
        r"C:\Users\r02an25\OneDrive - University of Aberdeen\Documents\Reports\Floating\Python_QGIS_DB\Layouts\qgis_export_catenary_dea.csv"
    )
    print(f"[INFO] Reverting to fallback: {csv_path}")
else:
  print("[ERROR] No MOORING_CSV_PATH environment variable was received.")
  csv_path = (
      r"C:\Users\r02an25\OneDrive - University of Aberdeen\Documents\Reports\Floating\Python_QGIS_DB\Layouts\qgis_export_catenary_dea.csv"
  )
  print(f"[INFO] Reverting to fallback: {csv_path}")

output_dir = os.path.join(project_root, project_name, "Layouts").replace("\\", "/")
os.makedirs(output_dir, exist_ok=True)
csv_filename_base = os.path.splitext(os.path.basename(csv_path))[0]
group_name = f"Layout = {csv_filename_base}"

root = QgsProject.instance().layerTreeRoot()
existing_group = root.findGroup(group_name)
if existing_group:
  root.removeChildNode(existing_group)
layer_group = root.insertGroup(0, group_name)

# Keep the Z field in the URI so all 3-D coordinate values are imported.
target_crs = "EPSG:3857"
uri = (
    f"file:///{csv_path}?encoding=UTF-8&delimiter=,&xField=X_Coord&"
    f"yField=Y_Coord&zField=Z_Coord&crs={target_crs}"
)
csv_layer = QgsVectorLayer(uri, "simulation_csv_data", "delimitedtext")

if not csv_layer.isValid():
  print("[ERROR] Failed to load CSV file into QGIS.")
else:
  print(
      "[SUCCESS] CSV loaded successfully in EPSG:3857. "
      f"Group created and placed at top: '{group_name}'"
  )

  # These match Main.py's export schema and retain the pre-existing layer
  # names, path construction, colours and individual bridle-branch grouping.
  categories = {
      "Turbine_Markers": {
          "filter": "Feature_Type = 'Turbine'", "to_path": False,
      },
      "Mooring_Radius": {
          "filter": "Feature_Type = 'Boundary_Radius' AND Sub_Type = 'Mooring_Radius_Circle'",
          "to_path": True, "close": True,
      },
      "DDP_Radius": {
          "filter": "Feature_Type = 'Boundary_Radius' AND Sub_Type = 'DDP_Radius_Circle'",
          "to_path": True, "close": True,
      },
      "TDP_Radius": {
          "filter": "Feature_Type = 'Boundary_Radius' AND Sub_Type = 'TDP_Radius_Circle'",
          "to_path": True, "close": True,
      },
      "Mooring_Line": {
          "filter": "Feature_Type = 'Mooring_Line'", "to_path": True, "close": False,
      },
      "Mooring_Joints": {
          "filter": "Feature_Type = 'Mooring_Joint'", "to_path": False,
      },
      "Anchor_Body": {
          "filter": "Feature_Type = 'Anchor_Body'", "to_path": True, "close": True,
      },
      "Anchor_Padeye": {
          "filter": "Feature_Type = 'Anchor_Component' AND Sub_Type = 'Padeye'",
          "to_path": False,
      },
      "Anchor_DDP": {
          "filter": "Feature_Type = 'Anchor_Component' AND Sub_Type = 'DDP'",
          "to_path": False,
      },
      "Anchor_TDP": {
          "filter": "Feature_Type = 'Anchor_Component' AND Sub_Type = 'TDP'",
          "to_path": False,
      },
      "Anchor_Fairlead": {
          "filter": "Feature_Type = 'Anchor_Component' AND Sub_Type = 'Fairlead'",
          "to_path": False,
      },
      "Farm_Perimeter": {
          "filter": "Feature_Type = 'Farm_Perimeter'", "to_path": True, "close": True,
      },
  }

  gpkg_filename = f"{csv_filename_base}.gpkg"
  gpkg_path = os.path.join(output_dir, gpkg_filename).replace("\\", "/")

  # Remove old loaded copies of this GeoPackage before replacing it.
  layers_to_remove = []
  for layer_id, layer in QgsProject.instance().mapLayers().items():
    source = layer.source() if hasattr(layer, "source") else ""
    if gpkg_filename.lower() in source.replace("\\", "/").lower():
      layers_to_remove.append(layer_id)
  if layers_to_remove:
    QgsProject.instance().removeMapLayers(layers_to_remove)
  gc.collect()

  if os.path.exists(gpkg_path):
    try:
      os.remove(gpkg_path)
      print(f"[INFO] Cleared existing GeoPackage: {gpkg_path}")
    except Exception as exc:
      timestamp = time.strftime("%H%M%S")
      gpkg_filename = f"{csv_filename_base}_{timestamp}.gpkg"
      gpkg_path = os.path.join(output_dir, gpkg_filename).replace("\\", "/")
      print(
          "[WARNING] GeoPackage locked by another process; redirecting output "
          f"to: {gpkg_filename} ({exc})"
      )

  transform_context = QgsProject.instance().transformContext()
  is_first_layer = True
  for layer_key, config in categories.items():
    extracted_layer = processing.run(
        "native:extractbyexpression",
        {
            "INPUT": csv_layer,
            "EXPRESSION": config["filter"],
            "OUTPUT": "memory:",
        },
    )["OUTPUT"]
    if not extracted_layer or extracted_layer.featureCount() == 0:
      continue

    export_source = extracted_layer
    if config["to_path"]:
      if layer_key == "Mooring_Line":
        # Includes Sub_Type so each bridle a/b branch remains its own path.
        group_expression = "concat(Turbine_ID, '_', Line_Heading_Deg, '_', Sub_Type)"
      elif layer_key == "Anchor_Body":
        group_expression = "concat(Turbine_ID, '_', Line_Heading_Deg)"
      else:
        group_expression = "Turbine_ID"
      export_source = processing.run(
          "native:pointstopath",
          {
              "INPUT": extracted_layer,
              "CLOSE_PATH": config["close"],
              "GROUP_EXPRESSION": group_expression,
              # Preserve Main.py's authoritative CSV insertion order.
              "ORDER_EXPRESSION": "",
              "NATURAL_SORT": False,
              "OUTPUT": "memory:",
          },
      ).get("OUTPUT")
      if layer_key == "Farm_Perimeter" and export_source is not None:
        export_source = processing.run(
            "native:polygonize",
            {
                "INPUT": export_source,
                "KEEP_FIELDS": True,
                "OUTPUT": "memory:",
            },
        )["OUTPUT"]

    if not export_source or not export_source.isValid():
      print(f"[ERROR] Failed to generate geometry for {layer_key}.")
      continue

    options = QgsVectorFileWriter.SaveVectorOptions()
    options.driverName = "GPKG"
    options.fileEncoding = "UTF-8"
    options.layerName = layer_key
    options.actionOnExistingFile = (
        QgsVectorFileWriter.CreateOrOverwriteFile
        if is_first_layer else QgsVectorFileWriter.CreateOrOverwriteLayer
    )
    is_first_layer = False
    error = QgsVectorFileWriter.writeAsVectorFormatV3(
        export_source, gpkg_path, transform_context, options
    )
    if error[0] != QgsVectorFileWriter.NoError:
      print(f"[ERROR] Failed to export layer {layer_key}: {error}")
      continue

    new_layer = QgsVectorLayer(f"{gpkg_path}|layername={layer_key}", layer_key, "ogr")
    if not new_layer.isValid():
      print(f"[ERROR] Could not load newly created layer for {layer_key}.")
      continue

    if layer_key == "Turbine_Markers":
      symbol = QgsMarkerSymbol.createSimple({"name": "triangle", "size": "5", "outline_color": "black"})
    elif layer_key == "Anchor_Padeye":
      symbol = QgsMarkerSymbol.createSimple({"name": "circle", "size": "3", "outline_color": "black"})
    elif layer_key == "Mooring_Joints":
      symbol = QgsMarkerSymbol.createSimple({"name": "diamond", "size": "4", "outline_color": "black"})
    elif layer_key in ("Anchor_DDP", "Anchor_TDP"):
      symbol = QgsMarkerSymbol.createSimple({"name": "square", "size": "3", "outline_color": "black"})
    elif layer_key == "Anchor_Fairlead":
      symbol = QgsMarkerSymbol.createSimple({"name": "circle", "size": "2.5", "outline_color": "black"})
    elif layer_key in ("Mooring_Radius", "DDP_Radius", "TDP_Radius"):
      symbol = QgsLineSymbol.createSimple({"line_style": "dash", "line_width": "0.4"})
    else:
      symbol = QgsLineSymbol.createSimple({"line_width": "0.8"})

    if layer_key == "Mooring_Line":
      color_expression = "coalesce(Color_Hex, '#1f2937')"
    elif layer_key == "Mooring_Radius":
      color_expression = "coalesce(Color_Hex, '#2563eb')"
    elif layer_key == "DDP_Radius":
      color_expression = "coalesce(Color_Hex, '#059669')"
    elif layer_key == "TDP_Radius":
      color_expression = "coalesce(Color_Hex, '#d97706')"
    else:
      color_expression = "coalesce(Color_Hex, '#4b5563')"
    color_property = QgsProperty.fromExpression(color_expression)
    geometry_type = new_layer.geometryType()
    if geometry_type == QgsWkbTypes.PointGeometry:
      symbol.symbolLayer(0).setDataDefinedProperty(
          QgsSymbolLayer.PropertyFillColor, color_property
      )
    elif geometry_type == QgsWkbTypes.LineGeometry:
      symbol.symbolLayer(0).setDataDefinedProperty(
          QgsSymbolLayer.PropertyStrokeColor, color_property
      )
    elif geometry_type == QgsWkbTypes.PolygonGeometry:
      symbol.symbolLayer(0).setDataDefinedProperty(
          QgsSymbolLayer.PropertyFillColor, color_property
      )

    new_layer.renderer().setSymbol(symbol)
    QgsProject.instance().addMapLayer(new_layer, False)
    layer_group.addLayer(new_layer)

  gc.collect()
  print(
      f"\n[INFO] Coordinate layers imported under top group '{group_name}'. "
      "Use Main.py's Unified Dashboard for Seismic Volume Calculation."
  )

import csv
import glob
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import tkinter as tkinter_module
import unicodedata
from tkinter import filedialog

# Force an interactive Matplotlib backend.  The dashboard uses Matplotlib
# Button/TextBox widgets, which do not receive mouse events on non-interactive
# backends such as Agg.  This must be selected BEFORE importing pyplot.
import matplotlib
try:
  if matplotlib.get_backend().lower() in ("agg", "module://matplotlib_inline.backend_inline"):
    matplotlib.use("TkAgg", force=True)
except Exception:
  # TkAgg will be selected by Matplotlib normally; keep the import robust.
  pass

import matplotlib.backend_bases
import matplotlib.patches as patches
import matplotlib.pyplot as plt
from matplotlib.widgets import Button, TextBox, Slider
from matplotlib.ticker import ScalarFormatter, MaxNLocator
import moorpy.Catenary as mc
import numpy as np
from scipy.optimize import least_squares
import xlwings as xw


def _ensure_geo_dependencies():
  """Load raster/CRS libraries only when a raster workflow actually needs them."""
  global rasterio, Transformer
  if "rasterio" not in globals():
    import rasterio as _rasterio
    rasterio = _rasterio
  if "Transformer" not in globals():
    from pyproj import Transformer as _Transformer
    Transformer = _Transformer


# =========================================================================
# NAVIA COLOUR RAMP
# =========================================================================
# Embedded from the user-supplied QGIS Navia.xml style.  The QGIS style
# defines a continuous RGB gradient from dark blue through cyan/green,
# orange and pale cream.  Keeping the stops here means the application does
# not depend on the XML file being present on another computer.
def create_navia_colormap():
  stops = [
    (0.0000000, (0.01176471, 0.07450980, 0.15294118)),
    (0.0039216, (0.01568627, 0.08235294, 0.16078431)),
    (0.0078431, (0.01568627, 0.08627451, 0.16862745)),
    (0.0117650, (0.01568627, 0.09019608, 0.17254902)),
    (0.0156860, (0.01568627, 0.09411765, 0.18039216)),
    (0.0196080, (0.01568627, 0.09803922, 0.18823529)),
    (0.0235290, (0.01960784, 0.10196078, 0.19607843)),
    (0.0274510, (0.01960784, 0.10588235, 0.20392157)),
    (0.0313730, (0.01960784, 0.10980392, 0.21176471)),
    (0.0352940, (0.01960784, 0.11372549, 0.21960784)),
    (0.0392160, (0.01960784, 0.11764706, 0.22745098)),
    (0.0431370, (0.01960784, 0.12549020, 0.23529412)),
    (0.0470590, (0.01960784, 0.12941176, 0.24313725)),
    (0.0509800, (0.01960784, 0.13333333, 0.25098039)),
    (0.0549020, (0.01960784, 0.13725490, 0.25882353)),
    (0.0588240, (0.01960784, 0.14117647, 0.26666667)),
    (0.0627450, (0.01960784, 0.14509804, 0.27450980)),
    (0.0666670, (0.01960784, 0.14901961, 0.28235294)),
    (0.0705880, (0.02352941, 0.15686275, 0.29019608)),
    (0.0745100, (0.02352941, 0.16078431, 0.29803922)),
    (0.0784310, (0.02352941, 0.16470588, 0.30588235)),
    (0.0823530, (0.02352941, 0.16862745, 0.31372549)),
    (0.0862750, (0.02352941, 0.17254902, 0.32156863)),
    (0.0901960, (0.02352941, 0.18039216, 0.32941176)),
    (0.0941180, (0.02352941, 0.18431373, 0.33725490)),
    (0.0980390, (0.02352941, 0.18823529, 0.34509804)),
    (0.1019600, (0.02352941, 0.19215686, 0.35294118)),
    (0.1058800, (0.02745098, 0.19607843, 0.36078431)),
    (0.1098000, (0.02745098, 0.20392157, 0.36862745)),
    (0.1137300, (0.02745098, 0.20784314, 0.37647059)),
    (0.1176500, (0.02745098, 0.21176471, 0.38431373)),
    (0.1215700, (0.02745098, 0.21568627, 0.39215686)),
    (0.1254900, (0.02745098, 0.22352941, 0.40000000)),
    (0.1294100, (0.03137255, 0.22745098, 0.40784314)),
    (0.1333300, (0.03137255, 0.23137255, 0.41568627)),
    (0.1372500, (0.03137255, 0.23529412, 0.41960784)),
    (0.1411800, (0.03137255, 0.24313725, 0.42745098)),
    (0.1451000, (0.03529412, 0.24705882, 0.43529412)),
    (0.1490200, (0.03529412, 0.25098039, 0.44313725)),
    (0.1529400, (0.03921569, 0.25490196, 0.45098039)),
    (0.1568600, (0.03921569, 0.26274510, 0.45490196)),
    (0.1607800, (0.04313725, 0.26666667, 0.46274510)),
    (0.1647100, (0.04313725, 0.27058824, 0.47058824)),
    (0.1686300, (0.04313725, 0.27450980, 0.47450980)),
    (0.1725500, (0.04705882, 0.28235294, 0.48235294)),
    (0.1764700, (0.05098039, 0.28627451, 0.49019608)),
    (0.1803900, (0.05098039, 0.29019608, 0.49411765)),
    (0.1843100, (0.05490196, 0.29803922, 0.50196078)),
    (0.1882400, (0.05490196, 0.30196078, 0.50588235)),
    (0.1921600, (0.05882353, 0.30588235, 0.50980392)),
    (0.1960800, (0.06274510, 0.30980392, 0.51764706)),
    (0.2000000, (0.06274510, 0.31764706, 0.52156863)),
    (0.2039200, (0.06666667, 0.32156863, 0.52549020)),
    (0.2078400, (0.07058824, 0.32549020, 0.52941176)),
    (0.2117600, (0.07450980, 0.32941176, 0.53333333)),
    (0.2156900, (0.07450980, 0.33725490, 0.53725490)),
    (0.2196100, (0.07843137, 0.34117647, 0.54117647)),
    (0.2235300, (0.08235294, 0.34509804, 0.54509804)),
    (0.2274500, (0.08627451, 0.34901961, 0.54901961)),
    (0.2313700, (0.09019608, 0.35294118, 0.55294118)),
    (0.2352900, (0.09019608, 0.36078431, 0.55294118)),
    (0.2392200, (0.09411765, 0.36470588, 0.55686275)),
    (0.2431400, (0.09803922, 0.36862745, 0.55686275)),
    (0.2470600, (0.10196078, 0.37254902, 0.56078431)),
    (0.2509800, (0.10588235, 0.37647059, 0.56078431)),
    (0.2549000, (0.10588235, 0.38039216, 0.56470588)),
    (0.2588200, (0.10980392, 0.38431373, 0.56470588)),
    (0.2627500, (0.11372549, 0.38823529, 0.56470588)),
    (0.2666700, (0.11764706, 0.39215686, 0.56862745)),
    (0.2705900, (0.11764706, 0.39607843, 0.56862745)),
    (0.2745100, (0.12156863, 0.40000000, 0.56862745)),
    (0.2784300, (0.12549020, 0.40392157, 0.56862745)),
    (0.2823500, (0.12549020, 0.40784314, 0.56862745)),
    (0.2862700, (0.12941176, 0.41176471, 0.56862745)),
    (0.2902000, (0.13333333, 0.41568627, 0.56862745)),
    (0.2941200, (0.13333333, 0.41960784, 0.56862745)),
    (0.2980400, (0.13725490, 0.42352941, 0.56862745)),
    (0.3019600, (0.14117647, 0.42352941, 0.56862745)),
    (0.3058800, (0.14117647, 0.42745098, 0.56862745)),
    (0.3098000, (0.14509804, 0.43137255, 0.56470588)),
    (0.3137300, (0.14901961, 0.43529412, 0.56470588)),
    (0.3176500, (0.14901961, 0.43529412, 0.56470588)),
    (0.3215700, (0.15294118, 0.43921569, 0.56470588)),
    (0.3254900, (0.15294118, 0.44313725, 0.56078431)),
    (0.3294100, (0.15686275, 0.44313725, 0.56078431)),
    (0.3333300, (0.15686275, 0.44705882, 0.56078431)),
    (0.3372500, (0.16078431, 0.45098039, 0.56078431)),
    (0.3411800, (0.16470588, 0.45098039, 0.55686275)),
    (0.3451000, (0.16470588, 0.45490196, 0.55686275)),
    (0.3490200, (0.16862745, 0.45882353, 0.55686275)),
    (0.3529400, (0.16862745, 0.45882353, 0.55294118)),
    (0.3568600, (0.17254902, 0.46274510, 0.55294118)),
    (0.3607800, (0.17254902, 0.46274510, 0.55294118)),
    (0.3647100, (0.17647059, 0.46666667, 0.54901961)),
    (0.3686300, (0.17647059, 0.47058824, 0.54901961)),
    (0.3725500, (0.18039216, 0.47058824, 0.54901961)),
    (0.3764700, (0.18431373, 0.47450980, 0.54509804)),
    (0.3803900, (0.18431373, 0.47450980, 0.54509804)),
    (0.3843100, (0.18823529, 0.47843137, 0.54509804)),
    (0.3882400, (0.18823529, 0.47843137, 0.54117647)),
    (0.3921600, (0.19215686, 0.48235294, 0.54117647)),
    (0.3960800, (0.19215686, 0.48627451, 0.54117647)),
    (0.4000000, (0.19607843, 0.48627451, 0.53725490)),
    (0.4039200, (0.19607843, 0.49019608, 0.53725490)),
    (0.4078400, (0.20000000, 0.49019608, 0.53725490)),
    (0.4117600, (0.20392157, 0.49411765, 0.53333333)),
    (0.4156900, (0.20392157, 0.49411765, 0.53333333)),
    (0.4196100, (0.20784314, 0.49803922, 0.53333333)),
    (0.4235300, (0.20784314, 0.49803922, 0.52941176)),
    (0.4274500, (0.21176471, 0.50196078, 0.52941176)),
    (0.4313700, (0.21176471, 0.50196078, 0.52549020)),
    (0.4352900, (0.21568627, 0.50588235, 0.52549020)),
    (0.4392200, (0.21568627, 0.50588235, 0.52549020)),
    (0.4431400, (0.21960784, 0.50980392, 0.52156863)),
    (0.4470600, (0.22352941, 0.51372549, 0.52156863)),
    (0.4509800, (0.22352941, 0.51372549, 0.52156863)),
    (0.4549000, (0.22745098, 0.51764706, 0.51764706)),
    (0.4588200, (0.22745098, 0.51764706, 0.51764706)),
    (0.4627500, (0.23137255, 0.52156863, 0.51764706)),
    (0.4666700, (0.23529412, 0.52156863, 0.51372549)),
    (0.4705900, (0.23529412, 0.52549020, 0.51372549)),
    (0.4745100, (0.23921569, 0.52549020, 0.50980392)),
    (0.4784300, (0.23921569, 0.52941176, 0.50980392)),
    (0.4823500, (0.24313725, 0.53333333, 0.50980392)),
    (0.4862700, (0.24705882, 0.53333333, 0.50588235)),
    (0.4902000, (0.24705882, 0.53725490, 0.50588235)),
    (0.4941200, (0.25098039, 0.53725490, 0.50588235)),
    (0.4980400, (0.25098039, 0.54117647, 0.50196078)),
    (0.5019600, (0.25490196, 0.54117647, 0.50196078)),
    (0.5058800, (0.25882353, 0.54509804, 0.49803922)),
    (0.5098000, (0.25882353, 0.54901961, 0.49803922)),
    (0.5137300, (0.26274510, 0.54901961, 0.49803922)),
    (0.5176500, (0.26666667, 0.55294118, 0.49411765)),
    (0.5215700, (0.26666667, 0.55294118, 0.49411765)),
    (0.5254900, (0.27058824, 0.55686275, 0.49019608)),
    (0.5294100, (0.27450980, 0.56078431, 0.49019608)),
    (0.5333300, (0.27843137, 0.56078431, 0.49019608)),
    (0.5372500, (0.27843137, 0.56470588, 0.48627451)),
    (0.5411800, (0.28235294, 0.56862745, 0.48627451)),
    (0.5451000, (0.28627451, 0.56862745, 0.48235294)),
    (0.5490200, (0.28627451, 0.57254902, 0.48235294)),
    (0.5529400, (0.29019608, 0.57647059, 0.47843137)),
    (0.5568600, (0.29411765, 0.57647059, 0.47843137)),
    (0.5607800, (0.29803922, 0.58039216, 0.47843137)),
    (0.5647100, (0.29803922, 0.58431373, 0.47450980)),
    (0.5686300, (0.30196078, 0.58431373, 0.47450980)),
    (0.5725500, (0.30588235, 0.58823529, 0.47058824)),
    (0.5764700, (0.30980392, 0.59215686, 0.47058824)),
    (0.5803900, (0.31372549, 0.59607843, 0.46666667)),
    (0.5843100, (0.31372549, 0.59607843, 0.46666667)),
    (0.5882400, (0.31764706, 0.60000000, 0.46274510)),
    (0.5921600, (0.32156863, 0.60392157, 0.46274510)),
    (0.5960800, (0.32549020, 0.60784314, 0.45882353)),
    (0.6000000, (0.32941176, 0.60784314, 0.45882353)),
    (0.6039200, (0.33333333, 0.61176471, 0.45490196)),
    (0.6078400, (0.33725490, 0.61568627, 0.45490196)),
    (0.6117600, (0.34117647, 0.61960784, 0.45098039)),
    (0.6156900, (0.34117647, 0.62352941, 0.45098039)),
    (0.6196100, (0.34509804, 0.62745098, 0.44705882)),
    (0.6235300, (0.34901961, 0.62745098, 0.44705882)),
    (0.6274500, (0.35294118, 0.63137255, 0.44313725)),
    (0.6313700, (0.35686275, 0.63529412, 0.44313725)),
    (0.6352900, (0.36078431, 0.63921569, 0.43921569)),
    (0.6392200, (0.36470588, 0.64313725, 0.43921569)),
    (0.6431400, (0.36862745, 0.64705882, 0.43529412)),
    (0.6470600, (0.37254902, 0.65098039, 0.43529412)),
    (0.6509800, (0.37647059, 0.65490196, 0.43137255)),
    (0.6549000, (0.38431373, 0.65882353, 0.43137255)),
    (0.6588200, (0.38823529, 0.66274510, 0.42745098)),
    (0.6627500, (0.39215686, 0.66666667, 0.42745098)),
    (0.6666700, (0.39607843, 0.67058824, 0.42352941)),
    (0.6705900, (0.40000000, 0.67450980, 0.42352941)),
    (0.6745100, (0.40392157, 0.67843137, 0.41960784)),
    (0.6784300, (0.41176471, 0.68235294, 0.41960784)),
    (0.6823500, (0.41568627, 0.68627451, 0.41568627)),
    (0.6862700, (0.41960784, 0.69019608, 0.41568627)),
    (0.6902000, (0.42745098, 0.69411765, 0.41176471)),
    (0.6941200, (0.43137255, 0.69803922, 0.41176471)),
    (0.6980400, (0.43921569, 0.70196078, 0.41176471)),
    (0.7019600, (0.44313725, 0.70588235, 0.41176471)),
    (0.7058800, (0.45098039, 0.70980392, 0.40784314)),
    (0.7098000, (0.45490196, 0.71372549, 0.40784314)),
    (0.7137300, (0.46274510, 0.71764706, 0.40784314)),
    (0.7176500, (0.47058824, 0.72549020, 0.40784314)),
    (0.7215700, (0.47450980, 0.72941176, 0.40784314)),
    (0.7254900, (0.48235294, 0.73333333, 0.40784314)),
    (0.7294100, (0.49019608, 0.73725490, 0.40784314)),
    (0.7333300, (0.49803922, 0.74117647, 0.40784314)),
    (0.7372500, (0.50588235, 0.74901961, 0.40784314)),
    (0.7411800, (0.51372549, 0.75294118, 0.40784314)),
    (0.7451000, (0.52156863, 0.75686275, 0.41176471)),
    (0.7490200, (0.52941176, 0.76078431, 0.41176471)),
    (0.7529400, (0.53725490, 0.76470588, 0.41568627)),
    (0.7568600, (0.54901961, 0.77254902, 0.41568627)),
    (0.7607800, (0.55686275, 0.77647059, 0.41960784)),
    (0.7647100, (0.56470588, 0.78039216, 0.42352941)),
    (0.7686300, (0.57647059, 0.78431373, 0.42745098)),
    (0.7725500, (0.58431373, 0.79215686, 0.43137255)),
    (0.7764700, (0.59607843, 0.79607843, 0.43529412)),
    (0.7803900, (0.60392157, 0.80000000, 0.43921569)),
    (0.7843100, (0.61176471, 0.80392157, 0.44705882)),
    (0.7882400, (0.62352941, 0.80784314, 0.45098039)),
    (0.7921600, (0.63529412, 0.81176471, 0.45490196)),
    (0.7960800, (0.64313725, 0.81960784, 0.46274510)),
    (0.8000000, (0.65490196, 0.82352941, 0.47058824)),
    (0.8039200, (0.66274510, 0.82745098, 0.47450980)),
    (0.8078400, (0.67450980, 0.83137255, 0.48235294)),
    (0.8117600, (0.68235294, 0.83529412, 0.49019608)),
    (0.8156900, (0.69411765, 0.83921569, 0.49803922)),
    (0.8196100, (0.70196078, 0.84313725, 0.50588235)),
    (0.8235300, (0.70980392, 0.84705882, 0.51372549)),
    (0.8274500, (0.72156863, 0.85098039, 0.52156863)),
    (0.8313700, (0.72941176, 0.85490196, 0.52941176)),
    (0.8352900, (0.73725490, 0.85882353, 0.53725490)),
    (0.8392200, (0.74901961, 0.86274510, 0.54509804)),
    (0.8431400, (0.75686275, 0.86274510, 0.55294118)),
    (0.8470600, (0.76470588, 0.86666667, 0.56470588)),
    (0.8509800, (0.77254902, 0.87058824, 0.57254902)),
    (0.8549000, (0.78039216, 0.87450980, 0.58039216)),
    (0.8588200, (0.78823529, 0.87843137, 0.58823529)),
    (0.8627500, (0.79607843, 0.87843137, 0.59607843)),
    (0.8666700, (0.80392157, 0.88235294, 0.60784314)),
    (0.8705900, (0.81176471, 0.88627451, 0.61568627)),
    (0.8745100, (0.81960784, 0.89019608, 0.62352941)),
    (0.8784300, (0.82745098, 0.89019608, 0.63137255)),
    (0.8823500, (0.83529412, 0.89411765, 0.63921569)),
    (0.8862700, (0.84313725, 0.89803922, 0.64705882)),
    (0.8902000, (0.84705882, 0.89803922, 0.65882353)),
    (0.8941200, (0.85490196, 0.90196078, 0.66666667)),
    (0.8980400, (0.86274510, 0.90588235, 0.67450980)),
    (0.9019600, (0.86666667, 0.90588235, 0.68235294)),
    (0.9058800, (0.87450980, 0.90980392, 0.69019608)),
    (0.9098000, (0.87843137, 0.90980392, 0.69803922)),
    (0.9137300, (0.88627451, 0.91372549, 0.70588235)),
    (0.9176500, (0.89019608, 0.91764706, 0.71372549)),
    (0.9215700, (0.89803922, 0.91764706, 0.72156863)),
    (0.9254900, (0.90196078, 0.92156863, 0.72941176)),
    (0.9294100, (0.90980392, 0.92156863, 0.73725490)),
    (0.9333300, (0.91372549, 0.92549020, 0.74509804)),
    (0.9372500, (0.91764706, 0.92941176, 0.74901961)),
    (0.9411800, (0.92156863, 0.92941176, 0.75686275)),
    (0.9451000, (0.92941176, 0.93333333, 0.76470588)),
    (0.9490200, (0.93333333, 0.93333333, 0.77254902)),
    (0.9529400, (0.93725490, 0.93725490, 0.77647059)),
    (0.9568600, (0.94117647, 0.93725490, 0.78431373)),
    (0.9607800, (0.94509804, 0.94117647, 0.79215686)),
    (0.9647100, (0.95294118, 0.94117647, 0.79607843)),
    (0.9686300, (0.95686275, 0.94509804, 0.80392157)),
    (0.9725500, (0.96078431, 0.94509804, 0.81176471)),
    (0.9764700, (0.96470588, 0.94901961, 0.81568627)),
    (0.9803900, (0.96862745, 0.94901961, 0.82352941)),
    (0.9843100, (0.97254902, 0.95294118, 0.82745098)),
    (0.9882400, (0.97647059, 0.95294118, 0.83529412)),
    (0.9921600, (0.98039216, 0.95294118, 0.83921569)),
    (0.9960800, (0.98431373, 0.95686275, 0.84313725)),
    (1.0000000, (0.98823529, 0.95686275, 0.85098039))
  ]
  return matplotlib.colors.LinearSegmentedColormap.from_list("navia", stops, N=256)


NAVIA_CMAP = create_navia_colormap()

# Shared UI palette based on the Failure Criteria report styling.
UI_NAVY = "#243746"
UI_BG = "#eef1f4"
UI_PANEL = "#f7f9fb"
UI_BORDER = "#d0d0d0"
UI_ACCENT = "#2f8f9d"
UI_ACCENT_HOVER = "#3aa6b5"
UI_PASS = "#287d46"
UI_FAIL = "#b33a3a"
UI_WARN = "#d9a441"
UI_TEXT = "#2c3e50"

# The workbook is chosen on the welcome screen.  Keeping the selected path in
# one place ensures any later Excel read (for example, from the setup screen)
# uses the same design database instead of whichever workbook happens to be
# active in Excel.
DEFAULT_EXCEL_WORKBOOK_PATH = (
    r"C:\Users\r02an25\OneDrive - University of Aberdeen\Documents\Reports\Floating\Floating_Offshore_Wind_Mooring&Anchoring_Design_Parameters.xlsx"
)
_SELECTED_EXCEL_WORKBOOK_PATH = None

# The farm-location picker and the raster-aware mooring checks must use the
# same definition of a suitable water depth.  The target is the design water
# depth read from Excel (E5); H30 supplies the +/- percentage threshold.
# This value is only a safe fallback when H30 is unavailable or invalid.
RASTER_DEPTH_TOLERANCE_FRACTION = 0.10
RASTER_DEPTH_BAND_EPSILON_M = 1e-6

# =========================================================================
# PATCH: FIX MATPLOTLIB WIDGET CRASHES (Python 3.14 / Matplotlib Compatibility)
# =========================================================================
if not hasattr(matplotlib.backend_bases.ResizeEvent, "inaxes"):
  matplotlib.backend_bases.ResizeEvent.inaxes = None
if not hasattr(matplotlib.backend_bases.ResizeEvent, "x"):
  matplotlib.backend_bases.ResizeEvent.x = 0
if not hasattr(matplotlib.backend_bases.ResizeEvent, "y"):
  matplotlib.backend_bases.ResizeEvent.y = 0
if not hasattr(matplotlib.backend_bases.ResizeEvent, "button"):
  matplotlib.backend_bases.ResizeEvent.button = None


# =========================================================================
# ESCAPE KEY EVENT HANDLER
# =========================================================================
def on_key_press(event):
  """Event handler to terminate the script when the Escape key is pressed."""
  if event.key == "escape":
    print(
        "\n[INFO] 'Escape' key pressed. Closing all windows and exiting"
        " program..."
    )
    plt.close("all")
    sys.exit(0)


_MACOS_OPEN_FILE_SCRIPT = """
on run argv
  set dialogTitle to item 1 of argv
  set selectedFile to choose file with prompt dialogTitle
  return POSIX path of selectedFile
end run
"""

_MACOS_SAVE_FILE_SCRIPT = """
on run argv
  set dialogTitle to item 1 of argv
  set suggestedName to item 2 of argv
  set selectedFile to choose file name with prompt dialogTitle default name suggestedName
  return POSIX path of selectedFile
end run
"""


def _run_macos_file_panel(script, *arguments):
  """Run a standard macOS Open/Save panel through ``osascript``.

  ``None`` means the native panel could not start, allowing a Tk fallback;
  an empty string is a normal user cancellation.  Arguments are supplied to
  AppleScript separately, rather than inserted into its source, so filenames
  containing quotes, spaces or Unicode cannot break the panel command.
  """
  try:
    result = subprocess.run(
        [
            "/usr/bin/osascript", "-l", "AppleScript", "-e", script,
            *[str(argument) for argument in arguments],
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
  except OSError as exc:
    print(f"[WARNING] macOS native file dialog could not start: {exc}")
    return None

  if result.returncode == 0:
    return (result.stdout or "").rstrip("\r\n")

  details = f"{result.stderr or ''}\n{result.stdout or ''}".strip()
  # Standard Additions returns -128 when the user presses Cancel.
  if "-128" in details or "user canceled" in details.lower():
    return ""
  print(
      f"[WARNING] macOS native file dialog failed: "
      f"{details or result.returncode}"
  )
  return None


def _show_macos_native_file_dialog(dialog_kind, title, initialfile="", defaultextension=""):
  """Use macOS's native Open/Save panel without crossing GUI toolkits."""
  if dialog_kind != "save":
    return _run_macos_file_panel(_MACOS_OPEN_FILE_SCRIPT, title)

  suggested_name = os.path.basename(str(initialfile or "untitled"))
  extension = str(defaultextension or "")
  if extension and not extension.startswith("."):
    extension = f".{extension}"
  if extension and not os.path.splitext(suggested_name)[1]:
    suggested_name += extension

  selected_path = _run_macos_file_panel(
      _MACOS_SAVE_FILE_SCRIPT, title, suggested_name
  )
  if selected_path and extension and not os.path.splitext(selected_path)[1]:
    selected_path += extension
  return selected_path


def _show_file_dialog(parent_figure, dialog_function, title, **dialog_options):
  """Show a Tk file dialog on Windows and other non-macOS platforms."""
  parent_window = None
  try:
    candidate = parent_figure.canvas.manager.window if parent_figure else None
    if isinstance(candidate, tkinter_module.Misc):
      parent_window = candidate
  except Exception:
    pass

  temporary_root = None
  topmost_enabled = False
  try:
    if parent_window is None:
      # Needed for the macOS Matplotlib backend and any other non-Tk backend.
      # Passing this root explicitly avoids a stale tkinter default root.
      temporary_root = tkinter_module.Tk()
      temporary_root.withdraw()
      try:
        temporary_root.update_idletasks()
      except Exception:
        pass
      parent_window = temporary_root

    # -topmost is useful on Windows, but is not consistently supported by
    # Tk/macOS and can prevent the native panel from appearing there.
    if sys.platform.startswith("win"):
      try:
        parent_window.attributes("-topmost", True)
        topmost_enabled = True
      except Exception:
        pass

    return dialog_function(
        parent=parent_window, title=title, **dialog_options
    )
  except Exception as exc:
    print(f"[WARNING] Could not open '{title}' file dialog: {exc}")
    return ""
  finally:
    if topmost_enabled:
      try:
        parent_window.attributes("-topmost", False)
      except Exception:
        pass
    if temporary_root is not None:
      try:
        temporary_root.destroy()
      except Exception:
        pass
    elif parent_window is not None:
      try:
        parent_window.focus_force()
      except Exception:
        pass


def _ask_for_file(parent_figure, title, filetypes):
  """Open a platform-native file chooser for an existing file."""
  if sys.platform == "darwin":
    native_path = _show_macos_native_file_dialog("open", title)
    if native_path is not None:
      return native_path
  return _show_file_dialog(
      parent_figure,
      filedialog.askopenfilename,
      title,
      filetypes=filetypes,
  )


def _ask_to_save_file(parent_figure, title, **dialog_options):
  """Open a platform-native Save As dialog."""
  if sys.platform == "darwin":
    native_path = _show_macos_native_file_dialog(
        "save",
        title,
        initialfile=dialog_options.get("initialfile", ""),
        defaultextension=dialog_options.get("defaultextension", ""),
    )
    if native_path is not None:
      return native_path
  return _show_file_dialog(
      parent_figure,
      filedialog.asksaveasfilename,
      title,
      **dialog_options,
  )


# =========================================================================
# HELPER: LAT/LON & EPSG:3857 (WEB MERCATOR) CONVERTERS
# =========================================================================
def latlon_to_epsg3857(lat, lon):
  """Converts WGS84 Latitude and Longitude (degrees) to EPSG:3857 (meters)."""
  R = 6378137.0
  x = R * np.radians(lon)
  y = R * np.log(np.tan(np.pi / 4.0 + np.radians(lat) / 2.0))
  return x, y


def open_qgis_and_import(csv_path, project_path=None, clip_rasters=True):
  """Launch QGIS to import the exported coordinate layers into a project.

  ``clip_rasters`` remains in the signature for compatibility with older
  callers, but QGIS no longer receives or processes raster/seismic work.
  """
  qgis_exe = find_qgis_executable()

  if qgis_exe and "qgis-bin.exe" in qgis_exe.lower():
    qgis_exe = qgis_exe.replace("qgis-bin.exe", "qgis.exe")

  # Resolve the importer location robustly.  In some interactive/embedded
  # Python launchers (notably when Main.py is executed through an IDE or
  # interactive console), __file__ is not defined.  Prefer it when available,
  # then fall back to the executed script path, the CSV location, and CWD.
  _candidate_dirs = []

  _this_file = globals().get("__file__")
  if _this_file:
    _candidate_dirs.append(os.path.dirname(os.path.abspath(_this_file)))

  try:
    _argv0 = sys.argv[0] if sys.argv else ""
    if _argv0 and os.path.isfile(_argv0):
      _candidate_dirs.append(os.path.dirname(os.path.abspath(_argv0)))
  except Exception:
    pass

  try:
    _csv_abs = os.path.abspath(csv_path)
    _csv_dir = os.path.dirname(_csv_abs)
    # CSVs are normally saved in Layouts beneath the Main.py project folder.
    _candidate_dirs.extend([_csv_dir, os.path.dirname(_csv_dir)])
  except Exception:
    pass

  _candidate_dirs.append(os.getcwd())

  # Remove duplicates while preserving priority order.
  _seen_dirs = set()
  _candidate_dirs = [
    d for d in _candidate_dirs
    if d and not (d in _seen_dirs or _seen_dirs.add(d))
  ]

  script_path = next(
    (os.path.join(d, "QGIS_Mooring_Importing.py")
     for d in _candidate_dirs
     if os.path.exists(os.path.join(d, "QGIS_Mooring_Importing.py"))),
    None,
  )

  if script_path is None:
    print("\n[ERROR] QGIS_Mooring_Importing.py could not be located.")
    print("[INFO] Searched:")
    for _d in _candidate_dirs:
      print(f"  - {_d}")
    return

  if not qgis_exe or not os.path.exists(qgis_exe):
    print("\n[ERROR] Could not find QGIS executable!")
    return

  qgis_exe = qgis_exe.replace("\\", "/")
  script_path = script_path.replace("\\", "/")
  csv_path = csv_path.replace("\\", "/")

  env = os.environ.copy()
  env["MOORING_CSV_PATH"] = csv_path

  cmd = [qgis_exe]

  if project_path and os.path.exists(project_path):
    project_path = project_path.replace("\\", "/")
    cmd.extend(["--project", project_path])
    print(f"\n[INFO] Target QGIS Project Layout: {project_path}")

  cmd.extend(["--code", script_path])

  print(f"[INFO] Launching QGIS: {qgis_exe}")
  print(f"[INFO] Executing script via --code: {script_path}")
  
  subprocess.Popen(cmd, env=env)


# =========================================================================
# WELCOME SCREEN / EXCEL DATABASE SELECTOR
# =========================================================================
def _normalise_excel_workbook_path(excel_path):
  """Return a stable absolute form of a user-provided workbook path."""
  expanded_path = os.path.expanduser(str(excel_path).strip())
  resolved_path = os.path.realpath(os.path.abspath(expanded_path))
  return os.path.normcase(unicodedata.normalize("NFC", resolved_path))


def _set_selected_excel_workbook_path(excel_path):
  """Persist the workbook chosen at startup for the rest of this session."""
  global _SELECTED_EXCEL_WORKBOOK_PATH
  _SELECTED_EXCEL_WORKBOOK_PATH = _normalise_excel_workbook_path(excel_path)


def launch_excel_welcome_screen(initial_excel_path=""):
  """Show the first screen and return the accepted Excel database path."""
  fig = plt.figure(figsize=(12, 6.5), facecolor=UI_BG)
  fig.canvas.mpl_connect("key_press_event", on_key_press)
  try:
    fig.canvas.manager.set_window_title("Nurdins Mooring Solution - Welcome")
  except Exception:
    pass

  fig.patches.append(
      patches.Rectangle(
          (0.0, 0.86), 1.0, 0.14, transform=fig.transFigure,
          facecolor=UI_NAVY, edgecolor="none", zorder=-1,
      )
  )
  fig.text(
      0.08, 0.91, "WELCOME", ha="left", va="center",
      fontsize=18, fontweight="bold", color="white",
  )
  fig.text(
      0.50, 0.82,
      "Start a mooring design assessment by selecting its Excel design database.",
      ha="center", va="center", fontsize=10, color=UI_TEXT,
  )

  panel = fig.add_axes([0.08, 0.21, 0.84, 0.49])
  panel.set_facecolor("white")
  panel.set_xticks([])
  panel.set_yticks([])
  for spine in panel.spines.values():
    spine.set_color(UI_BORDER)

  fig.text(
      0.11, 0.62, "Excel Design Database", fontsize=11,
      fontweight="bold", color=UI_TEXT,
  )
  fig.text(
      0.11, 0.58,
      "Choose the workbook containing the 'Mooring_Python' worksheet.",
      fontsize=9, color="#555555",
  )
  ax_excel = fig.add_axes([0.11, 0.48, 0.66, 0.058])
  txt_excel = TextBox(
      ax_excel, "", initial=initial_excel_path or "", textalignment="left"
  )
  ax_excel_btn = fig.add_axes([0.79, 0.48, 0.10, 0.058])
  btn_excel = Button(
      ax_excel_btn, "Browse...", color="#e9ecef", hovercolor=UI_ACCENT_HOVER
  )
  fig.text(
      0.11, 0.39,
      "Supported files: .xlsx, .xlsm, .xlsb, and .xls.  The selected workbook "
      "remains the source of all dashboard calculations for this session.",
      fontsize=9, color="#555555", va="top",
  )

  status_text = fig.text(
      0.11, 0.285, "", fontsize=9, color=UI_FAIL, va="center"
  )
  ax_continue = fig.add_axes([0.34, 0.10, 0.32, 0.065])
  btn_continue = Button(
      ax_continue, "Open Dashboard", color="lightgray", hovercolor="lightgray"
  )
  for widget_axis in (ax_excel, ax_excel_btn, ax_continue):
    widget_axis.set_zorder(panel.get_zorder() + 10)

  state = {"excel": initial_excel_path or "", "accepted_path": None}
  supported_extensions = {".xlsx", ".xlsm", ".xlsb", ".xls"}

  def _entered_excel_path():
    # Accept a quoted path copied from Explorer without making the displayed
    # path unexpectedly change while the user is editing it.
    return txt_excel.text.strip().strip('"')

  def update_state(show_error=False):
    state["excel"] = _entered_excel_path()
    extension = os.path.splitext(state["excel"])[1].lower()
    resolved_path = (
        _normalise_excel_workbook_path(state["excel"])
        if state["excel"] else ""
    )
    valid = (
        bool(state["excel"])
        and extension in supported_extensions
        and os.path.isfile(resolved_path)
    )
    btn_continue.color = UI_ACCENT if valid else "lightgray"
    btn_continue.hovercolor = UI_ACCENT_HOVER if valid else "lightgray"
    btn_continue.ax.set_facecolor(btn_continue.color)
    if show_error:
      if not state["excel"]:
        status_text.set_text("Please enter or browse to an Excel design database.")
      elif extension not in supported_extensions:
        status_text.set_text("Please select an Excel workbook (.xlsx, .xlsm, .xlsb, or .xls).")
      else:
        status_text.set_text("The selected Excel workbook could not be found.")
    elif valid:
      status_text.set_text("")
    fig.canvas.draw_idle()
    return valid

  txt_excel.on_submit(lambda text: update_state())

  def browse_excel(event=None):
    path = _ask_for_file(
        fig,
        title="Select Excel Design Database",
        filetypes=[
            ("Excel Workbooks", ("*.xlsx", "*.xlsm", "*.xlsb", "*.xls")),
            ("All Files (*.*)", "*.*"),
        ],
    )
    if path:
      state["excel"] = path
      txt_excel.set_val(path)
      update_state()

  def continue_clicked(event=None):
    if not update_state(show_error=True):
      return

    excel_path = _normalise_excel_workbook_path(state["excel"])
    try:
      # Open and check the required sheet before leaving the welcome screen,
      # so an invalid workbook does not start the expensive solver workflow.
      get_excel_sheet("Mooring_Python", excel_path=excel_path)
    except Exception as exc:
      details = str(exc).splitlines()
      detail = details[0] if details else type(exc).__name__
      status_text.set_text(
          "Could not open the 'Mooring_Python' worksheet: "
          f"{detail}"
      )
      status_text.set_color(UI_FAIL)
      fig.canvas.draw_idle()
      return

    # Let the outer ``main`` call continue only after this window's Tk event
    # loop has returned.  Starting the project screen inside this callback can
    # leave file dialogs attached to a closing/stale Tk window.
    state["accepted_path"] = excel_path
    plt.close(fig)

  btn_excel.on_clicked(browse_excel)
  btn_continue.on_clicked(continue_clicked)
  # Keep interactive widgets alive on TkAgg/Windows even if Matplotlib is in
  # a non-blocking/interactive mode and this function's local frame returns.
  fig._welcome_widgets = [txt_excel, btn_excel, btn_continue]
  update_state()

  try:
    manager = plt.get_current_backend_manager()
    manager.full_screen_toggle()
  except Exception:
    try:
      fig.canvas.manager.window.state("zoomed")
    except Exception:
      pass

  plt.show()
  return state["accepted_path"]


# =========================================================================
# ENHANCED INTERACTIVE MAP LOCATION PICKER (SHOWS REAL FARM GEOMETRY)
# =========================================================================
def launch_project_setup_screen(
    results_store,
    system_flags,
    anchor_db,
    project_path="",
    raster_path="",
    water_depth=None,
    depth_threshold=None,
):
  """Collect the QGIS project and raster after the workbook is selected."""
  fig = plt.figure(figsize=(12, 6.5), facecolor=UI_BG)
  fig.canvas.mpl_connect("key_press_event", on_key_press)
  try:
    fig.canvas.manager.set_window_title("Nurdins Mooring Solution - Project Setup")
  except Exception:
    pass

  fig.patches.append(patches.Rectangle((0.0, 0.86), 1.0, 0.14, transform=fig.transFigure,
                                        facecolor=UI_NAVY, edgecolor="none", zorder=-1))
  fig.text(
      0.08, 0.91, "PROJECT SETUP", ha="left", va="center",
      fontsize=18, fontweight="bold", color="white"
  )
  fig.text(
      0.50, 0.82,
      "Select the QGIS project and raster used for the wind-farm location picker.",
      ha="center", va="center", fontsize=10
  )

  panel = fig.add_axes([0.08, 0.18, 0.84, 0.54])
  panel.set_facecolor("white")
  panel.set_xticks([])
  panel.set_yticks([])
  for spine in panel.spines.values():
    spine.set_color(UI_BORDER)

  fig.text(0.11, 0.64, "1. QGIS Project", fontsize=10, fontweight="bold", color=UI_TEXT)
  ax_qgis = fig.add_axes([0.11, 0.56, 0.66, 0.055])
  txt_qgis = TextBox(ax_qgis, "", initial=project_path or "", textalignment="left")
  ax_qgis_btn = fig.add_axes([0.79, 0.56, 0.10, 0.055])
  btn_qgis = Button(ax_qgis_btn, "Browse...", color="#e9ecef", hovercolor=UI_ACCENT_HOVER)

  fig.text(0.11, 0.47, "2. Raster File", fontsize=10, fontweight="bold", color=UI_TEXT)
  ax_raster = fig.add_axes([0.11, 0.39, 0.66, 0.055])
  txt_raster = TextBox(ax_raster, "", initial=raster_path or "", textalignment="left")
  ax_raster_btn = fig.add_axes([0.79, 0.39, 0.10, 0.055])
  btn_raster = Button(ax_raster_btn, "Browse...", color="#e9ecef", hovercolor=UI_ACCENT_HOVER)

  fig.text(
      0.11, 0.30,
      "The raster is displayed on the next screen. Only one point can be selected;\n"
      "the wind-farm perimeter is generated later from the selected mooring system.",
      fontsize=9, va="top", color="#555555"
  )

  ax_continue = fig.add_axes([0.34, 0.08, 0.32, 0.065])
  btn_continue = Button(
      ax_continue, "Continue to Location Picker",
      color="lightgray", hovercolor=UI_ACCENT_HOVER
  )
  # The white panel is itself an Axes.  Put all interactive child axes above
  # it explicitly so it cannot receive the click intended for a Browse button.
  for widget_axis in (
      ax_qgis, ax_qgis_btn, ax_raster, ax_raster_btn, ax_continue,
  ):
    widget_axis.set_zorder(panel.get_zorder() + 10)

  state = {"project": project_path or "", "raster": raster_path or ""}

  def update_state():
    state["project"] = txt_qgis.text.strip()
    state["raster"] = txt_raster.text.strip()
    valid = (
        bool(state["project"])
        and bool(state["raster"])
        and os.path.isfile(state["project"])
        and os.path.isfile(state["raster"])
    )
    btn_continue.color = UI_ACCENT if valid else "lightgray"
    btn_continue.hovercolor = UI_ACCENT_HOVER if valid else "lightgray"
    btn_continue.ax.set_facecolor(btn_continue.color)
    fig.canvas.draw_idle()
    return valid

  txt_qgis.on_submit(lambda text: update_state())
  txt_raster.on_submit(lambda text: update_state())

  def browse_qgis(event=None):
    path = _ask_for_file(
        fig,
        title="Select QGIS Project File",
        filetypes=[
            ("QGIS Projects", ("*.qgz", "*.qgs")),
            ("All Files (*.*)", "*.*"),
        ],
    )
    if path:
      state["project"] = path
      txt_qgis.set_val(path)
      update_state()

  def browse_raster(event=None):
    path = _ask_for_file(
        fig,
        title="Select Raster File",
        filetypes=[
            ("Raster Files", ("*.tif", "*.tiff", "*.img", "*.vrt")),
            ("All Files (*.*)", "*.*"),
        ],
    )
    if path:
      state["raster"] = path
      txt_raster.set_val(path)
      update_state()

  def continue_clicked(event=None):
    if not update_state():
      print("\n[WARNING] Please select an existing QGIS project and raster file.")
      return
    project_path = state["project"]
    raster_path = state["raster"]
    
    # Reuse the values read from the selected welcome-screen workbook.  This
    # avoids a later COM lookup accidentally reading a different active Excel
    # workbook.  Retain the fallback for direct callers of this setup screen.
    if water_depth is None or depth_threshold is None:
      try:
        sheet = get_excel_sheet()
        selected_water_depth = excel_cell(sheet, "water_depth", default=100.0)
        selected_depth_threshold = read_raster_depth_threshold(sheet)
      except Exception as e:
        print(f"\n[WARNING] Could not read water depth/H30 threshold from Excel: {e}. Defaulting to 100.0m and +/-10%.")
        selected_water_depth = 100.0
        selected_depth_threshold = RASTER_DEPTH_TOLERANCE_FRACTION
    else:
      selected_water_depth = water_depth
      selected_depth_threshold = depth_threshold
      
    plt.close(fig)
    open_raster_location_picker(
        raster_path,
        lambda lat, lon: (
            update_result_locations(results_store, lat, lon),
            launch_unified_dashboard(
                results_store,
                system_flags,
                anchor_db,
                project_path=project_path,
                raster_path=raster_path,
            ),
        ),
        water_depth=selected_water_depth,
        depth_threshold=selected_depth_threshold,
        on_back=lambda: launch_project_setup_screen(
            results_store, system_flags, anchor_db,
            project_path=project_path,
            raster_path=raster_path,
            water_depth=selected_water_depth,
            depth_threshold=selected_depth_threshold,
        ),
    )

  btn_qgis.on_clicked(browse_qgis)
  btn_raster.on_clicked(browse_raster)
  btn_continue.on_clicked(continue_clicked)
  # Retain widgets and their callbacks explicitly.  This is especially
  # important on the Windows TkAgg backend when ``plt.show`` is non-blocking.
  fig._project_setup_widgets = [
      txt_qgis, txt_raster, btn_qgis, btn_raster, btn_continue,
  ]

  try:
    manager = plt.get_current_backend_manager()
    manager.full_screen_toggle()
  except Exception:
    try:
      fig.canvas.manager.window.state("zoomed")
    except Exception:
      pass

  plt.show()

def update_result_locations(results_store, center_lat, center_lon):
  """Apply the user-selected wind-farm centre to every stored simulation result."""
  for system_results in results_store.values():
    for data in system_results.values():
      if isinstance(data, dict):
        data["center_lat"] = center_lat
        data["center_lon"] = center_lon


def open_raster_location_picker(
    raster_path,
    on_confirm,
    water_depth=100.0,
    depth_threshold=RASTER_DEPTH_TOLERANCE_FRACTION,
    initial_lat=None,
    initial_lon=None,
    on_back=None,
):
  """Display the user-selected raster and allow exactly one centre point to be picked."""
  _ensure_geo_dependencies()
  depth_band = get_raster_depth_band(water_depth, depth_threshold)
  if depth_band is None:
    print("[ERROR] Excel E5 water depth or H30 depth threshold is invalid.")
    return
  try:
    src = rasterio.open(raster_path)
  except Exception as e:
    print(f"\n[ERROR] Could not open raster file '{raster_path}': {e}")
    return

  with src:
    if src.crs is None:
      print("[WARNING] Raster has no CRS. Assuming EPSG:4326 (WGS84).")
      raster_crs = "EPSG:4326"
    else:
      raster_crs = src.crs

    # Read a display-sized version so very large bathymetry/terrain rasters do not
    # consume excessive memory in the picker.
    max_dim = 1800
    scale = min(1.0, max_dim / max(src.width, src.height))
    out_h = max(1, int(src.height * scale))
    out_w = max(1, int(src.width * scale))

    # The location picker is a bathymetry/terrain map, so use the supplied
    # Navia scientific colour ramp for the first raster band.
    display_arr = src.read(1, out_shape=(out_h, out_w), masked=True)

    bounds = src.bounds
    transform_to_wgs84 = Transformer.from_crs(raster_crs, "EPSG:4326", always_xy=True)

  fig, ax = plt.subplots(figsize=(12, 8.5), facecolor=UI_BG)
  fig.canvas.mpl_connect("key_press_event", on_key_press)
  try:
    fig.canvas.manager.set_window_title("Nurdins Mooring Solution - Location Picker")
  except Exception:
    pass
  fig.subplots_adjust(top=0.86, bottom=0.15, left=0.07, right=0.97)
  fig.patches.append(patches.Rectangle((0.0, 0.90), 1.0, 0.10, transform=fig.transFigure,
                                        facecolor=UI_NAVY, edgecolor="none", zorder=-1))
  fig.text(0.07, 0.945, "SELECT WIND FARM CENTRE", ha="left", va="center",
           fontsize=16, fontweight="bold", color="white")
  fig.text(0.07, 0.915, "Choose a point inside the permitted water-depth band", ha="left", va="center",
           fontsize=9, color="#dce6ed")

  # Display the selected raster using the embedded Navia colour ramp.
  valid_values = np.asarray(display_arr.compressed() if np.ma.isMaskedArray(display_arr) else display_arr).ravel()
  valid_values = valid_values[np.isfinite(valid_values)]
  if valid_values.size:
    vmin = float(np.percentile(valid_values, 2.0))
    vmax = float(np.percentile(valid_values, 98.0))
    if not np.isfinite(vmin) or not np.isfinite(vmax) or vmax <= vmin:
      vmin = float(np.min(valid_values))
      vmax = float(np.max(valid_values))
  else:
    vmin, vmax = None, None

  ax.imshow(
      display_arr,
      extent=(bounds.left, bounds.right, bounds.bottom, bounds.top),
      origin="upper",
      cmap=NAVIA_CMAP,
      vmin=vmin,
      vmax=vmax,
    )

  # === CALCULATE EXCLUSION ZONE (E5 WATER DEPTH +/- H30) ===
  # Keep this exactly aligned with the turbine/anchor checks performed after
  # the farm layout has been generated.
  wd = depth_band["design_depth"]
  min_z = -depth_band["max_depth"]
  max_z = -depth_band["min_depth"]

  # Identify regions outside the valid depth window
  invalid_mask = (display_arr < min_z) | (display_arr > max_z)
  
  # Ensure NoData boundaries aren't heavily hashed unless they are actual data
  if np.ma.isMaskedArray(display_arr):
      invalid_mask = invalid_mask & ~display_arr.mask

  # Draw Red Overlay with Hashes
  if np.any(invalid_mask):
      x_grid = np.linspace(bounds.left, bounds.right, out_w)
      y_grid = np.linspace(bounds.bottom, bounds.top, out_h)
      
      # Flip the mask vertically so origin='upper' raster aligns with standard Y-axis for contourf
      # contourf accepts 1-D X/Y coordinate arrays, so avoid allocating two
      # full 1800 x 1800 mesh grids just for the overlay.  uint8 keeps the
      # temporary mask compact without changing the contour levels.
      invalid_mask_flipped = np.asarray(
          np.flipud(invalid_mask), dtype=np.uint8
      )
      
      # Base red layer (transparent)
      ax.contourf(x_grid, y_grid, invalid_mask_flipped, levels=[0.5, 1.5], colors=['red'], alpha=0.3, zorder=2)
      
      # Hatches
      old_hatch_color = matplotlib.rcParams.get('hatch.color', 'black')
      matplotlib.rcParams['hatch.color'] = 'darkred'
      ax.contourf(x_grid, y_grid, invalid_mask_flipped, levels=[0.5, 1.5], colors='none', hatches=['//'], zorder=3)
      matplotlib.rcParams['hatch.color'] = old_hatch_color

  ax.set_title(
      "Bathymetry / Water-Depth Selection\n"
      f"Allowed Depths: {min_z:.1f}m to {max_z:.1f}m "
      f"(Target: -{wd:.1f}m; H30: +/-{depth_band['tolerance_fraction'] * 100.0:.2f}%)",
      fontsize=11, fontweight="bold", color=UI_TEXT
  )
  ax.set_xlabel("Easting", fontsize=8.5, color=UI_TEXT, labelpad=7)
  ax.set_ylabel("Northing", fontsize=8.5, color=UI_TEXT, labelpad=7)
  ax.xaxis.set_major_locator(MaxNLocator(nbins=5))
  ax.yaxis.set_major_locator(MaxNLocator(nbins=5))
  xfmt = ScalarFormatter(useOffset=True, useMathText=False)
  yfmt = ScalarFormatter(useOffset=True, useMathText=False)
  xfmt.set_powerlimits((-3, 6))
  yfmt.set_powerlimits((-3, 6))
  ax.xaxis.set_major_formatter(xfmt)
  ax.yaxis.set_major_formatter(yfmt)
  ax.tick_params(axis="both", which="major", labelsize=8, colors=UI_TEXT, pad=3)
  for tick in ax.get_xticklabels():
    tick.set_rotation(0)
    tick.set_ha("center")
  ax.grid(True, linestyle=":", linewidth=0.7, alpha=0.30, color="#6c7a86")

  selected = {"x": None, "y": None, "lat": None, "lon": None}
  marker = None

  # Restore the previously selected centre when returning from the dashboard.
  if initial_lat is not None and initial_lon is not None:
    try:
      ix, iy = Transformer.from_crs("EPSG:4326", raster_crs, always_xy=True).transform(
          float(initial_lon), float(initial_lat)
      )
      if bounds.left <= ix <= bounds.right and bounds.bottom <= iy <= bounds.top:
        selected.update(x=float(ix), y=float(iy), lat=float(initial_lat), lon=float(initial_lon))
        marker, = ax.plot(
            selected["x"], selected["y"], "o", markersize=10,
            markerfacecolor="red", markeredgecolor="black", zorder=10
        )
    except Exception as e:
      print(f"[WARNING] Could not restore selected centre on raster: {e}")
  coord_label = (
      f"Selected centre: Latitude {selected['lat']:.6f}° | Longitude {selected['lon']:.6f}°"
      if selected["lat"] is not None else "No point selected"
  )
  coord_text = fig.text(
      0.50, 0.875, coord_label, ha="center", va="center",
      fontsize=10, fontweight="bold", color=UI_TEXT
  )

  ax_back = fig.add_axes([0.05, 0.035, 0.22, 0.065])
  btn_back = Button(ax_back, "← Back to Project Setup", color="#e9ecef", hovercolor=UI_ACCENT_HOVER)

  ax_confirm = fig.add_axes([0.33, 0.035, 0.34, 0.065])
  btn_confirm = Button(
      ax_confirm, "Confirm Centre Point",
      color="lightgray", hovercolor=UI_ACCENT_HOVER
  )

  def select_point(event):
    nonlocal marker
    if event.inaxes != ax or event.xdata is None or event.ydata is None:
      return

    # === VALIDATE SELECTED DEPTH ===
    col = int((event.xdata - bounds.left) / (bounds.right - bounds.left) * out_w)
    row = int((bounds.top - event.ydata) / (bounds.top - bounds.bottom) * out_h)
    
    if 0 <= col < out_w and 0 <= row < out_h:
        val = display_arr[row, col]
        
        if np.ma.is_masked(val) or np.isnan(val) or val < min_z or val > max_z:
            val_str = "NoData" if (np.ma.is_masked(val) or np.isnan(val)) else f"{val:.1f}"
            coord_text.set_text(f"Invalid area (Depth: {val_str}m). Allowed: {min_z:.1f}m to {max_z:.1f}m")
            coord_text.set_color("red")
            
            # Remove previous selection point if it exists
            if marker is not None:
                marker.remove()
                marker = None
                
            selected["lat"] = None
            selected["lon"] = None
            btn_confirm.color = "lightgray"
            btn_confirm.hovercolor = "lightgray"
            btn_confirm.ax.set_facecolor("lightgray")
            fig.canvas.draw_idle()
            return
            
    coord_text.set_color(UI_TEXT)

    selected["x"] = float(event.xdata)
    selected["y"] = float(event.ydata)
    lon, lat = transform_to_wgs84.transform(selected["x"], selected["y"])
    selected["lat"] = float(lat)
    selected["lon"] = float(lon)

    if marker is not None:
      marker.remove()
    marker, = ax.plot(
        selected["x"], selected["y"], "o",
        markersize=10, markerfacecolor="red", markeredgecolor="black",
        zorder=10,
    )
    coord_text.set_text(
        f"Selected centre: Latitude {lat:.6f}° | Longitude {lon:.6f}°"
    )
    btn_confirm.color = UI_ACCENT
    btn_confirm.hovercolor = UI_ACCENT_HOVER
    btn_confirm.ax.set_facecolor(UI_ACCENT)
    fig.canvas.draw_idle()

  def confirm_point(event):
    if selected["lat"] is None or selected["lon"] is None:
      print("\n[WARNING] Select one point on the raster before confirming.")
      return
    lat = selected["lat"]
    lon = selected["lon"]
    plt.close(fig)
    on_confirm(lat, lon)

  def back_to_setup(event=None):
    plt.close(fig)
    if on_back is not None:
      on_back()

  fig.canvas.mpl_connect("button_press_event", select_point)
  btn_confirm.on_clicked(confirm_point)
  btn_back.on_clicked(back_to_setup)
  fig._picker_widgets = [btn_confirm, btn_back]


  try:
    manager = plt.get_current_backend_manager()
    manager.full_screen_toggle()
  except Exception:
    pass

  plt.show()


# =========================================================================
# HELPER: SUBSURFACE GEOMETRY & ARC LENGTH CALCULATOR
# =========================================================================
def _validate_padeye_parameters(padeye_params):
  """Validate user-defined padeye angle/position inputs.

  Angles are degrees from horizontal. Position fractions are measured downwards
  from the TOP of cylindrical anchors: 0.0 = top, 1.0 = bottom.
  """
  defaults = {
      "DEA": {"angle_deg": 15.0},
      "Suction Pile": {"angle_deg": 15.0, "position_fraction": 0.50},
      "Driven Pile": {"angle_deg": 15.0, "position_fraction": 0.50},
      "Drilled Pile": {"angle_deg": 15.0, "position_fraction": 0.50},
  }
  params = dict(defaults)
  if padeye_params:
    for key, val in padeye_params.items():
      if key in params and isinstance(val, dict):
        params[key] = {**params[key], **val}

  for anchor, cfg in params.items():
    angle = float(cfg.get("angle_deg", 15.0))
    if not (0.0 < angle < 89.0):
      raise ValueError(f"{anchor} padeye angle must be between 0 and 89 degrees from horizontal; got {angle}.")
    if "position_fraction" in cfg:
      frac = float(cfg["position_fraction"])
      if not (0.0 <= frac <= 1.0):
        raise ValueError(f"{anchor} padeye position must be between 0.0 and 1.0; got {frac}.")
      cfg["position_fraction"] = frac
    cfg["angle_deg"] = angle
  return params


def get_subsurface_geometry(
    water_depth, anchor_db, primary_anc="DEA", padeye_params=None
):
  """Calculate the embedded/subsurface mooring geometry for Catenary/Semi-Taut.

  The padeye is fixed to the physical anchor geometry; changing its location
  does NOT translate the anchor. The line leaves the padeye at the user-specified
  angle from horizontal, then blends to horizontal at the DDP.
  """
  dims = anchor_db.get(
      primary_anc, {"depth": 8.0, "height": 3.0, "width": 4.0}
  )
  max_p_depth = float(dims.get("depth", 8.0))
  height = float(dims.get("height", 3.0))
  if max_p_depth <= 0.0 or height <= 0.0:
    raise ValueError(f"Invalid anchor geometry for {primary_anc}: depth={max_p_depth}, height={height}")

  params = _validate_padeye_parameters(padeye_params)
  cfg = params.get(primary_anc, {"angle_deg": 15.0, "position_fraction": 0.50})

  seabed_z = -float(water_depth)
  z_bottom = seabed_z - max_p_depth

  # Cylindrical anchors use `height` as their actual vertical length.
  # The position fraction is measured downwards from the top.
  cylindrical = any(k in str(primary_anc).upper() for k in ("SUCTION", "DRIVEN", "DRILLED", "PILE"))
  if cylindrical and primary_anc in ("Suction Pile", "Driven Pile", "Drilled Pile"):
    pos_frac = float(cfg.get("position_fraction", 0.50))
    z_top = z_bottom + height
    z_padeye = z_top - pos_frac * height
  else:
    # DEA/other non-cylindrical behaviour retains its existing mid-height
    # reference because no padeye position input is requested for DEA.
    z_padeye = z_bottom + (0.50 * height)

  embedment = seabed_z - z_padeye
  if embedment <= 0.0:
    raise ValueError(
        f"{primary_anc} padeye is not below the seabed: embedment={embedment:.3f} m. "
        "Check anchor penetration/height and padeye position."
    )

  angle_deg = float(cfg.get("angle_deg", 15.0))
  slope_padeye = np.tan(np.radians(angle_deg))
  if slope_padeye <= 0.0:
    raise ValueError(f"Invalid subsurface angle for {primary_anc}: {angle_deg} degrees")

  x_sub_span = (2.0 * embedment) / slope_padeye
  sub_x = np.linspace(0.0, x_sub_span, 100)
  A = -slope_padeye / (2.0 * x_sub_span) if x_sub_span > 0.0 else 0.0
  B = slope_padeye
  C = z_padeye
  sub_z = A * (sub_x**2) + B * sub_x + C

  dx = np.diff(sub_x)
  dz = np.diff(sub_z)
  L_sub = float(np.sum(np.sqrt(dx**2 + dz**2)))

  return sub_x, sub_z, L_sub, x_sub_span, z_padeye, angle_deg


def get_pile_anchor_vertical_geometry(
    water_depth, anchor_depth, anchor_height, padeye_fraction=0.25
):
  """Return fixed cylindrical anchor bottom/top/padeye elevations.

  Position fraction is measured downward from the TOP of the anchor.
  0.25 therefore places the padeye 25% down the anchor length.
  """
  seabed_z = -float(water_depth)
  z_bottom = seabed_z - float(anchor_depth)
  z_top = z_bottom + float(anchor_height)
  z_padeye = z_top - (float(padeye_fraction) * float(anchor_height))
  return z_bottom, z_top, z_padeye


# =========================================================================
# HELPER: MBL CALCULATORS AND DYNAMIC GRADE SELECTION
# =========================================================================
def calculate_chain_mbl(d_m, grade="R4"):
  d_mm = d_m * 1000.0
  c_mbl = {
      "R3": 0.0223,
      "R3S": 0.0249,
      "R4": 0.0274,
      "R4S": 0.0304,
      "R5": 0.0320,
  }
  c = c_mbl.get(grade, 0.0274)
  return c * (d_mm**2) * (44.0 - 0.08 * d_mm) * 1000.0


def select_optimal_chain_grade(d_m, T_max, min_fos=1.0):
  for grade in ["R3", "R3S", "R4", "R4S", "R5"]:
    mbl = calculate_chain_mbl(d_m, grade)
    fos = mbl / T_max if T_max > 0 else float("inf")
    if fos >= min_fos:
      return grade, mbl, fos
  mbl_r5 = calculate_chain_mbl(d_m, "R5")
  return "R5", mbl_r5, (mbl_r5 / T_max if T_max > 0 else 0.0)


def calculate_rope_mbl(d_m, material="Polyester"):
  d_mm = d_m * 1000.0
  c_mbl = {"Polyester": 0.33, "Nylon": 0.28, "HMPE": 0.65}
  return c_mbl.get(material, 0.33) * (d_mm**2) * 1000.0


# =========================================================================
# HELPER: ROBUST EXCEL CELL READING & ANCHOR DATABASE
# =========================================================================
# =========================================================================
# CENTRAL EXCEL INPUT CONFIGURATION
# =========================================================================
# EDIT EXCEL CELL LOCATIONS HERE ONLY.
#
# IMPORTANT:
# This section ONLY centralises Excel cell references. It does not change
# any engineering calculations, system logic, defaults, return values, or
# system-specific mooring inputs.
#
# If an input moves in Excel, change the cell/range on the RIGHT only.
# =========================================================================

EXCEL_CELLS = {
    # ---- Common / main inputs -------------------------------------------
    "water_depth": "E5",
    "number_of_turbines": "E7",
    "user_defined_buffer": "E8",
    "user_defined_radius": "E14",
    "fairlead_depth": "E20",

    # ---- Padeye inputs ---------------------------------------------------
    "dea_padeye_angle": "E24",
    "suction_padeye_angle": "E28",
    "suction_padeye_position": "E31",
    "driven_padeye_angle": "E35",
    "driven_padeye_position": "E38",
    "drilled_padeye_angle": "E42",
    "drilled_padeye_position": "E45",

    # ---- Farm / location inputs -----------------------------------------
    "farm_area": "H5",
    "buffer_zone": "E8",              # active user-defined buffer
    "center_latitude": "E50",
    "center_longitude": "E51",
    "taut_percentage": "H16",
    "triad": "H11",
    "inter_arm_angle": "H12",
    "angle_between_mooring_arms": "L22",
    "upper_equals_lower": "H19",
    "lower_joint_pos": "H21",
    "inverted_bridle": "H24",
    "inter_anchor_angle": "H25",
    "raster_depth_threshold": "H30",

    # ---- Catenary inputs -------------------------------------------------
    "catenary_anchor_radius": "L5",
    "catenary_line_length": "L6",
    "catenary_chain_diameter": "L7",
    "catenary_number_of_lines": "L8",

    # ---- Semi-taut inputs -----------------------------------------------
    "semi_taut_anchor_radius": "L11",
    "semi_taut_line_length": "L12",
    "semi_taut_rope_diameter": "L13",
    "semi_taut_chain_diameter": "L14",
    "semi_taut_number_of_lines": "L15",

    # ---- Taut inputs -----------------------------------------------------
    "taut_anchor_radius": "L18",
    "taut_line_diameter": "L19",
    "taut_number_of_lines": "L20",

    # ---- System execution switches --------------------------------------
    "run_catenary": "C5",
    "run_semi_taut": "C11",
    "run_taut": "C18",
}

EXCEL_RANGES = {
    # Anchor database
    "anchor_areas": "P5:P9",
    "anchor_depths": "P12:P16",
    "anchor_heights": "P19:P23",
    "anchor_widths": "P26:P30",
}


def excel_cell(sheet, name, default=0.0, cast_type=float):
    """Read a named Excel input from EXCEL_CELLS."""
    if name not in EXCEL_CELLS:
        raise KeyError(f"Excel input '{name}' is not defined in EXCEL_CELLS.")
    return get_cell_val(
        sheet,
        EXCEL_CELLS[name],
        default=default,
        cast_type=cast_type,
    )


def read_raster_depth_threshold(sheet, default=RASTER_DEPTH_TOLERANCE_FRACTION):
    """Read H30 and retain a safe percentage fallback for old workbooks."""
    raw_threshold = excel_cell(
        sheet,
        "raster_depth_threshold",
        default=default,
        cast_type=float,
    )
    if normalise_raster_depth_threshold(raw_threshold) is None:
        print(
            f"[WARNING] Invalid raster depth threshold in Excel H30 ({raw_threshold!r}); "
            f"using +/-{float(default) * 100.0:.1f}% instead."
        )
        return float(default)
    return raw_threshold


def excel_range(sheet, name):
    """Read a named Excel range from EXCEL_RANGES."""
    if name not in EXCEL_RANGES:
        raise KeyError(f"Excel range '{name}' is not defined in EXCEL_RANGES.")
    return sheet.range(EXCEL_RANGES[name]).value


def validate_mooring_angle_configuration(
    num_lines,
    triad,
    inter_arm_angle_deg,
    angle_between_mooring_arms_deg,
    tolerance_deg=1e-6,
):
  """Double-check the Excel-defined mooring-line angular arrangement.

  Triad=True:
      num_lines must divide into 3 equal arms.
      Within an arm, adjacent lines are separated by inter_arm_angle.
      Between arms, the gap from the last line of one arm to the first
      line of the next is angle_between_mooring_arms.

      Required closure:
          3 * ((lines_per_arm - 1) * inter_arm_angle
               + angle_between_mooring_arms) = 360°

  Triad=False:
      Every adjacent line is separated by angle_between_mooring_arms.

      Required closure:
          num_lines * angle_between_mooring_arms = 360°
  """
  n = int(num_lines)
  triad = bool(triad)
  inter = float(inter_arm_angle_deg)
  outer = float(angle_between_mooring_arms_deg)

  if n < 1:
    raise ValueError(f"Number of mooring lines must be >= 1; got {n}.")

  if not (0.0 < outer <= 360.0):
    raise ValueError(
        f"angle_between_mooring_arms must be > 0 and <= 360°; got {outer}."
    )

  if triad:
    if n % 3 != 0:
      raise ValueError(
          f"Triad=True requires the number of mooring lines to be divisible "
          f"by 3; got {n}."
      )

    lines_per_arm = n // 3

    if lines_per_arm > 1 and not (0.0 < inter < 180.0):
      raise ValueError(
          f"inter_arm_angle must be > 0 and < 180° when an arm contains "
          f"multiple lines; got {inter}°."
      )

    expected_total = (
        3.0 * outer
        if lines_per_arm == 1
        else 3.0 * ((lines_per_arm - 1) * inter + outer)
    )

    if not np.isclose(expected_total, 360.0, atol=tolerance_deg, rtol=0.0):
      raise ValueError(
          f"Invalid triad angles: 3 × (({lines_per_arm}-1) × "
          f"{inter:.6f}° + {outer:.6f}°) = {expected_total:.6f}°, "
          f"not 360°."
      )
  else:
    expected_total = n * outer
    if not np.isclose(expected_total, 360.0, atol=tolerance_deg, rtol=0.0):
      raise ValueError(
          f"Invalid non-triad angles: {n} × {outer:.6f}° = "
          f"{expected_total:.6f}°, not 360°."
      )

  return True


def generate_mooring_headings(
    num_lines,
    triad=False,
    inter_arm_angle_deg=0.0,
    angle_between_mooring_arms_deg=120.0,
):
  """Generate the actual mooring-line azimuths.

  Example:
      Triad=True, 6 lines, 10° inter-arm, 110° outer:
      [0°, 10°, 120°, 130°, 240°, 250°]

  Triad=False:
      headings are [0, outer, 2*outer, ...].
  """
  n = int(num_lines)
  triad = bool(triad)
  inter = float(inter_arm_angle_deg)
  outer = float(angle_between_mooring_arms_deg)

  validate_mooring_angle_configuration(n, triad, inter, outer)

  if not triad:
    headings = np.arange(n, dtype=float) * outer
  else:
    lines_per_arm = n // 3
    headings = []
    current = 0.0

    for _arm in range(3):
      for line_idx in range(lines_per_arm):
        headings.append(current)
        if line_idx < lines_per_arm - 1:
          current += inter
      current += outer

    headings = np.asarray(headings, dtype=float)

  return np.mod(headings, 360.0)




def validate_inverted_bridle_configuration(
    triad,
    upper_equals_lower,
    inverted_bridle,
    inter_anchor_angle_deg,
    lower_joint_pos,
):
  """Validate the Excel-controlled inverted-bridle configuration.

  An inverted bridle is deliberately incompatible with the triad azimuth
  arrangement because each original mooring heading must be split into one
  symmetric two-anchor branch pair.

  H19 controls whether the upper/lower section is treated as equal.  H21 is
  only required for the non-equal case and is a fraction measured from the
  fairlead down the suspended line (0 = fairlead, 1 = end of suspension).
  """
  if not bool(inverted_bridle):
    return True

  if bool(triad):
    raise ValueError(
        "Inverted Bridle cannot be enabled while Triad (H11) is TRUE. "
        "Set H11=FALSE before enabling H24."
    )

  angle = float(inter_anchor_angle_deg)
  if not np.isfinite(angle) or not (0.0 < angle < 180.0):
    raise ValueError(
        f"Inter_Anchor_Angle (H25) must be > 0° and < 180°; got {angle}°."
    )

  if not bool(upper_equals_lower):
    pos = float(lower_joint_pos)
    if not np.isfinite(pos) or not (0.0 < pos < 1.0):
      raise ValueError(
          f"Lower_Joint_Pos (H21) must be between 0 and 1 when H19=FALSE; got {pos}."
      )

  return True


def _profile_arc_length(x, z):
  x = np.asarray(x, dtype=float)
  z = np.asarray(z, dtype=float)
  if x.size < 2 or z.size != x.size:
    return np.array([], dtype=float)
  ds = np.hypot(np.diff(x), np.diff(z))
  return np.concatenate(([0.0], np.cumsum(ds)))


def _interpolate_profile_at_arc_fraction(x, z, fraction_from_start):
  """Interpolate a profile point using true 2-D arc length."""
  x = np.asarray(x, dtype=float)
  z = np.asarray(z, dtype=float)
  s = _profile_arc_length(x, z)
  if s.size < 2 or s[-1] <= 0.0:
    raise ValueError("Cannot locate inverted-bridle joint on an invalid profile.")
  target = float(np.clip(fraction_from_start, 0.0, 1.0)) * s[-1]
  return (
      float(np.interp(target, s, x)),
      float(np.interp(target, s, z)),
      float(s[-1]),
  )


def calculate_inverted_bridle_geometry(
    data, heading_deg, turbine_xy=None, raster_path=None, raster_sampler=None
):
  """Build the geometric inverted-Y branch arrangement without changing the solver.

  The existing solved mooring profile remains the trunk from fairlead to the
  lower joint.  At the lower joint the line is split into two equal azimuthal
  branches.  The two anchor padeyes are placed at the same radial anchor
  position and at +/- H25/2 around the original mooring heading.

  For Catenary the joint is located on the suspended TDP->fairlead section.
  For Taut it is located on the complete suspended profile.  Semi-Taut uses
  its existing lower chain/rope joint (rope_nodes[0]).
  """
  if not bool(data.get("inverted_bridle", False)):
    return None

  system_type = str(data.get("system_type", "Catenary"))
  triad = bool(data.get("triad", False))
  upper_equals_lower = bool(data.get("upper_equals_lower", True))
  lower_joint_pos = float(data.get("lower_joint_pos", 0.5))
  inter_anchor_angle = float(data.get("inter_anchor_angle", 30.0))
  validate_inverted_bridle_configuration(
      triad, upper_equals_lower, True, inter_anchor_angle, lower_joint_pos
  )

  x_plot = np.asarray(data.get("x_plot", []), dtype=float)
  z_plot = np.asarray(data.get("z_plot", []), dtype=float)
  sub_x = np.asarray(data.get("sub_x", []), dtype=float)
  sub_z = np.asarray(data.get("sub_z", []), dtype=float)
  rope_nodes = data.get("rope_nodes")
  X_td = data.get("X_td")

  # Semi-Taut always uses the solver's existing Bottom Chain / Rope Joint.
  # When H19=FALSE that existing joint has already been repositioned by H21
  # inside evaluate_semi_taut_anchor(), so do not create/interpolate a second
  # joint here.
  if system_type == "Semi-Taut" and rope_nodes is not None:
    joint_x, joint_z = float(rope_nodes[0]), float(rope_nodes[1])
    joint_source = (
        "existing Bottom Chain / Rope Joint (H21-controlled)"
        if not upper_equals_lower
        else "existing Bottom Chain / Rope Joint (H19=TRUE)"
    )
  else:
    if sub_x.size > 0 and sub_z.size == sub_x.size:
      full_x = np.concatenate((sub_x, x_plot[1:]))
      full_z = np.concatenate((sub_z, z_plot[1:]))
    else:
      full_x, full_z = x_plot, z_plot

    if full_x.size < 2 or full_z.size != full_x.size:
      raise ValueError("Inverted Bridle requires a valid solved mooring profile.")

    # H21 is measured from the FAIRLEAD toward the TDP for Catenary and
    # Semi-Taut when H19=FALSE. The grounded DDP->TDP section is excluded.
    # Taut has no grounded section, so its complete suspended profile is used.
    use_suspended_only = (
        system_type == "Catenary"
        or (system_type == "Semi-Taut" and not upper_equals_lower)
    )
    if use_suspended_only and X_td is not None:
      x_td = float(X_td)
      mask = full_x >= x_td - 1e-9
      susp_x = full_x[mask]
      susp_z = full_z[mask]
      if susp_x.size < 2:
        raise ValueError(
            "Inverted Bridle cannot locate the H21 joint on the suspended TDP->fairlead section."
        )
      joint_x, joint_z, suspended_arc = _interpolate_profile_at_arc_fraction(
          susp_x[::-1], susp_z[::-1], lower_joint_pos
      )
      joint_source = "H21 fraction of suspended TDP->fairlead section, measured from fairlead"
    else:
      joint_x, joint_z, suspended_arc = _interpolate_profile_at_arc_fraction(
          full_x[::-1], full_z[::-1], lower_joint_pos
      )
      joint_source = "H21 fraction of suspended line, measured from fairlead"

  if turbine_xy is None:
    turbine_xy = (0.0, 0.0)
  cx, cy = map(float, turbine_xy)
  heading = float(heading_deg)
  anchor_radius = float(data.get("anchor_radius", 0.0))
  anchor_width = float(data.get("anchor_width", 0.0))
  r_padeye = (
      anchor_radius - anchor_width / 2.0
      if system_type == "Taut"
      else anchor_radius - anchor_width
  )

  # Local solved x is measured radially inward from the anchor/padeye toward
  # the fairlead. Therefore the joint's global radial location is r_padeye-joint_x.
  joint_radial = r_padeye - joint_x
  h = np.radians(heading)
  joint_global = (
      cx + joint_radial * np.cos(h),
      cy + joint_radial * np.sin(h),
      joint_z,
  )

  half = 0.5 * inter_anchor_angle
  anchor_headings = [heading - half, heading + half]

  # Use the PADeye elevation from the actual solved local profile first.
  # This is critical for raster-depth re-evaluation: the local result may have a
  # different water depth from the Excel target depth. Recomputing the padeye
  # from the global data here can therefore make the bridle jump vertically.
  anc_name = str(data.get("primary_anc", "DEA")).strip()
  anc_upper = anc_name.upper()
  solver_padeye_z = data.get("padeye_z")
  try:
    solver_padeye_z = float(solver_padeye_z)
  except (TypeError, ValueError):
    solver_padeye_z = np.nan

  if system_type == "Taut":
    # Taut bridle padeyes are independently locked to the local raster seabed
    # at EACH physical branch anchor XY.  The anchor body extends downward
    # from that padeye; never reuse the turbine's global/Excel depth here.
    branch_padeye_z = np.nan
  elif np.isfinite(solver_padeye_z):
    # Catenary/Semi-Taut solver value. This is the same physical padeye used by
    # the subsurface geometry and the normal anchor rendering.
    branch_padeye_z = solver_padeye_z
  elif any(k in anc_upper for k in ("SUCTION", "DRIVEN", "DRILLED", "PILE")):
    frac = float((data.get("padeye_params") or {}).get(anc_name, {}).get("position_fraction", 0.50))
    _, _, branch_padeye_z = get_pile_anchor_vertical_geometry(
        float(data.get("water_depth", 0.0)),
        float(data.get("anchor_depth", 0.0)),
        float(data.get("anchor_height", 0.0)),
        padeye_fraction=frac,
    )
  else:
    branch_padeye_z = (
        float(sub_z[0]) if sub_z.size and np.isfinite(sub_z[0])
        else -float(data.get("water_depth", 0.0))
    )
  branch_padeye_z = float(branch_padeye_z)

  branch_anchors = []
  for branch_heading in anchor_headings:
    a = np.radians(branch_heading)
    px = cx + r_padeye * np.cos(a)
    py = cy + r_padeye * np.sin(a)
    if system_type == "Taut":
      geom = get_taut_anchor_seabed_geometry(
          raster_path, px, py,
          fallback_z=(-float(data.get("water_depth", 0.0))),
          anchor_height=float(data.get("anchor_height", 0.0)),
          raster_sampler=raster_sampler,
      )
      branch_z = geom["padeye_z"]
    else:
      branch_z = branch_padeye_z
    branch_anchors.append({
        "heading_deg": float(branch_heading % 360.0),
        "x": float(px),
        "y": float(py),
        "z": float(branch_z),
        "seabed_z": float(branch_z) if system_type == "Taut" else None,
        "anchor_top_z": float(branch_z) if system_type == "Taut" else None,
        "anchor_bottom_z": (float(branch_z) - float(data.get("anchor_height", 0.0))) if system_type == "Taut" else None,
    })

  # Branch lengths are the actual 3-D distances from Joint 1 to the two
  # anchor padeyes. They are equal by construction when the anchors are
  # symmetric about the original mooring heading.
  branch_lengths = []
  for a in branch_anchors:
    branch_lengths.append(float(np.linalg.norm(np.array([
        a["x"] - joint_global[0],
        a["y"] - joint_global[1],
        a["z"] - joint_global[2],
    ]))))

  return {
      "enabled": True,
      "joint": joint_global,
      "joint_local": (float(joint_x), float(joint_z)),
      "joint_source": joint_source,
      "anchor_headings_deg": anchor_headings,
      "anchors": branch_anchors,
      "branch_lengths": branch_lengths,
      "inter_anchor_angle_deg": inter_anchor_angle,
      "upper_equals_lower": upper_equals_lower,
      "lower_joint_pos": lower_joint_pos,
  }



def build_inverted_bridle_branch_profiles(data, bridle_geometry, ground_profile=None):
  """Build continuous anchor->lower-joint profiles for both inverted-bridle branches.

  This is a rendering transformation only: it uses the already-solved mooring
  profile and does not re-solve the Catenary/Semi-Taut/Taut system.

  Critical rule for Semi-Taut/Catenary:
    * Keep the COMPLETE solved anchor->fairlead profile as the source geometry.
    * If a raster-ground profile is supplied, replace ONLY DDP->TDP elevations.
    * Never replace the whole profile with the raster-only DDP->TDP array; doing
      that drops the suspended bottom-chain section between TDP and the Lower
      Joint and creates the visual break seen in the 2D bridle profile.
  """
  if not bridle_geometry or not bridle_geometry.get("enabled"):
    return []

  x_plot = np.asarray(data.get("x_plot", []), dtype=float)
  z_plot = np.asarray(data.get("z_plot", []), dtype=float)
  sub_x = np.asarray(data.get("sub_x", []), dtype=float)
  sub_z = np.asarray(data.get("sub_z", []), dtype=float)
  system_type = str(data.get("system_type", "Catenary"))

  # Build the complete ORIGINAL solved profile in anchor -> fairlead order.
  if sub_x.size >= 2 and sub_z.size == sub_x.size:
    full_x = np.concatenate((sub_x, x_plot[1:] if x_plot.size else np.array([])))
    full_z = np.concatenate((sub_z, z_plot[1:] if z_plot.size else np.array([])))
  else:
    full_x = x_plot.copy()
    full_z = z_plot.copy()

  finite = np.isfinite(full_x) & np.isfinite(full_z)
  full_x, full_z = full_x[finite], full_z[finite]
  if full_x.size < 2:
    return []

  order = np.argsort(full_x)
  full_x, full_z = full_x[order], full_z[order]
  keep = np.r_[True, np.diff(full_x) > 1e-10]
  full_x, full_z = full_x[keep], full_z[keep]

  joint_x, joint_z = map(float, bridle_geometry.get("joint_local", (np.nan, np.nan)))
  if not np.isfinite(joint_x) or not np.isfinite(joint_z):
    return []
  joint_x = float(np.clip(joint_x, full_x[0], full_x[-1]))

  # Base profile is ALWAYS the solver profile.  Only overwrite the actual
  # grounded DDP->TDP interval with raster elevations.
  profile_x = full_x.copy()
  profile_z = full_z.copy()
  if ground_profile is not None and system_type in ("Catenary", "Semi-Taut"):
    try:
      gx = np.asarray(ground_profile.get("x", []), dtype=float)
      gz = np.asarray(ground_profile.get("z", []), dtype=float)
      x_ddp = ground_profile.get("x_ddp")
      x_td = ground_profile.get("x_td")
      if gx.size >= 2 and gz.size == gx.size and x_ddp is not None and x_td is not None:
        good = np.isfinite(gx) & np.isfinite(gz)
        gx, gz = gx[good], gz[good]
        if gx.size >= 2:
          go = np.argsort(gx)
          gx, gz = gx[go], gz[go]
          lo = min(float(x_ddp), float(x_td))
          hi = max(float(x_ddp), float(x_td))
          ground_mask = (profile_x >= lo - 1e-9) & (profile_x <= hi + 1e-9)
          if np.any(ground_mask):
            profile_z[ground_mask] = np.interp(
                profile_x[ground_mask], gx, gz, left=gz[0], right=gz[-1]
            )
    except Exception:
      pass

  # Insert exact TDP/DDP samples into the rendered profile where possible so the
  # raster transition is represented on the same continuous curve as the solver.
  if ground_profile is not None:
    try:
      gx = np.asarray(ground_profile.get("x", []), dtype=float)
      gz = np.asarray(ground_profile.get("z", []), dtype=float)
      x_ddp = ground_profile.get("x_ddp")
      x_td = ground_profile.get("x_td")
      if gx.size >= 2 and gz.size == gx.size and x_ddp is not None and x_td is not None:
        inserts = [float(x_ddp), float(x_td)]
        for xx in inserts:
          if full_x[0] < xx < full_x[-1] and not np.any(np.isclose(profile_x, xx, atol=1e-9)):
            zz = float(np.interp(xx, gx, gz)) if gx.size else float(np.interp(xx, profile_x, profile_z))
            profile_x = np.append(profile_x, xx)
            profile_z = np.append(profile_z, zz)
        order = np.argsort(profile_x)
        profile_x, profile_z = profile_x[order], profile_z[order]
        keep = np.r_[True, np.diff(profile_x) > 1e-10]
        profile_x, profile_z = profile_x[keep], profile_z[keep]
    except Exception:
      pass

  # Keep only anchor -> Lower Joint, but retain the actual suspended profile
  # between TDP and the Lower Joint.
  mask = profile_x <= joint_x + 1e-9
  lx = profile_x[mask].copy()
  lz = profile_z[mask].copy()
  if lx.size < 2:
    lx = np.array([profile_x[0], joint_x], dtype=float)
    lz = np.array([profile_z[0], joint_z], dtype=float)
  else:
    if lx[0] > 1e-8:
      lx = np.insert(lx, 0, 0.0)
      lz = np.insert(lz, 0, float(profile_z[0]))
    if lx[-1] < joint_x - 1e-8:
      lx = np.append(lx, joint_x)
      lz = np.append(lz, float(np.interp(joint_x, profile_x, profile_z)))
    else:
      lx[-1] = joint_x
      lz[-1] = joint_z

  # The physical branch MUST begin at the physical padeye and finish at the
  # actual solver Lower Joint.
  anchors = bridle_geometry.get("anchors", [])
  if len(anchors) < 2:
    return []

  # Re-impose exact local endpoints after all insertion/sorting operations.
  lx[0] = 0.0
  lx[-1] = joint_x
  lz[-1] = joint_z

  # Make the start elevation come from each physical branch padeye.
  branches = []
  q = np.clip(
      (lx - float(lx[0])) / max(float(lx[-1] - lx[0]), 1e-12), 0.0, 1.0
  )
  cx, cy, _ = map(float, bridle_geometry["joint"])

  for idx, anchor in enumerate(anchors[:2], start=1):
    ax = float(anchor["x"])
    ay = float(anchor["y"])
    az = float(anchor["z"])

    bx = ax + q * (cx - ax)
    by = ay + q * (cy - ay)
    # Catenary and Semi-Taut branch endpoints are already taken from the same
    # validated solver profile as the normal (non-bridle) system.  Preserve
    # that curve exactly: an affine remap or end-smoothing here would alter its
    # natural tangent at the Lower Joint and create the visible "tendril" bend
    # when it joins the unchanged trunk.
    #
    # Taut is deliberately different.  Each physical bridle padeye is sampled
    # at its own raster XY position, so its start elevation can differ from the
    # centreline solver datum.  Only Taut therefore needs an endpoint-preserving
    # affine Z remap to connect that physical padeye to the common lower joint.
    lz = np.asarray(lz, dtype=float)
    if system_type != "Taut":
        _endpoint_tolerance = float(
            globals().get("GEOMETRY_ENDPOINT_TOLERANCE", 0.50)
        )
        if (
            not np.isfinite(lz[0])
            or not np.isfinite(lz[-1])
            or abs(float(lz[0]) - float(az)) > _endpoint_tolerance
            or abs(float(lz[-1]) - float(joint_z)) > _endpoint_tolerance
        ):
            raise ValueError(
                "Inverted Bridle endpoint differs from the normal solver "
                "profile; refusing to distort the Catenary/Semi-Taut curve."
            )
        bz = lz.copy()
    elif lz.size >= 2 and np.isfinite(lz[0]) and np.isfinite(lz[-1]):
        old_span = float(lz[-1] - lz[0])
        new_span = float(joint_z - az)
        if abs(old_span) > 1e-9:
            bz = float(az) + (lz - float(lz[0])) * (new_span / old_span)
        else:
            bz = np.linspace(float(az), float(joint_z), lz.size)
    else:
        bz = np.linspace(float(az), float(joint_z), max(len(lx), 2))

    bz = np.asarray(bz, dtype=float)

    bz[0] = float(az)
    bz[-1] = float(joint_z)

    # IMPORTANT FOR THE 2D PROFILE: keep the original solver radial x-coordinate.
    # H25 is a plan-view separation and must not change the horizontal scale of
    # the vertical profile. The 3-D branch geometry above remains unchanged.
    profile_x_2d = lx.copy()
    branch_joint_x_2d = float(joint_x)

    branches.append({
        "branch": idx,
        "x": bx,
        "y": by,
        "z": bz,
        "profile_x": lx.copy(),
        "profile_x_2d": profile_x_2d,
        "profile_z": bz.copy(),
        "joint_x_2d": branch_joint_x_2d,
        "anchor": dict(anchor),
        "system_type": system_type,
    })

  return branches


def validate_inverted_bridle_branch_profiles(bridle_geometry, tolerance=None):
  """Validate the rendered bridle endpoints against the normal solver geometry.

  Bridle rendering deliberately reuses the normal system's solved profile.
  Validate the two transformed branches after construction so a rendering
  change cannot silently detach a physical anchor or lower joint from that
  authoritative solution.
  """
  if tolerance is None:
    tolerance = float(globals().get("GEOMETRY_ENDPOINT_TOLERANCE", 0.50))
  tolerance = max(float(tolerance), 1e-9)
  if not bridle_geometry or not bridle_geometry.get("enabled"):
    return False, "Inverted Bridle geometry is not enabled."

  profiles = bridle_geometry.get("profiles", [])
  anchors = bridle_geometry.get("anchors", [])
  try:
    joint = np.asarray(bridle_geometry["joint"], dtype=float)
  except (KeyError, TypeError, ValueError):
    return False, "Inverted Bridle Lower Joint is invalid."
  if len(profiles) != 2 or len(anchors) < 2 or joint.size != 3 or not np.all(np.isfinite(joint)):
    return False, "Inverted Bridle must contain two finite physical branches."

  for branch_index, (branch, anchor) in enumerate(zip(profiles[:2], anchors[:2]), start=1):
    bx = np.asarray(branch.get("x", []), dtype=float)
    by = np.asarray(branch.get("y", []), dtype=float)
    bz = np.asarray(branch.get("z", []), dtype=float)
    px = _bridle_branch_2d_x(branch)
    if bx.size < 2 or by.size != bx.size or bz.size != bx.size or px.size != bx.size:
      return False, f"Inverted Bridle branch {branch_index} profile is incomplete."
    if not (np.all(np.isfinite(bx)) and np.all(np.isfinite(by)) and np.all(np.isfinite(bz))):
      return False, f"Inverted Bridle branch {branch_index} has non-finite coordinates."
    if np.any(np.diff(px) <= 1e-10):
      return False, f"Inverted Bridle branch {branch_index} is not monotonic in the 2-D profile."
    try:
      expected_anchor = np.asarray(
          [anchor["x"], anchor["y"], anchor["z"]], dtype=float
      )
    except (KeyError, TypeError, ValueError):
      return False, f"Inverted Bridle branch {branch_index} anchor is invalid."
    start_error = float(np.linalg.norm(np.array([bx[0], by[0], bz[0]]) - expected_anchor))
    end_error = float(np.linalg.norm(np.array([bx[-1], by[-1], bz[-1]]) - joint))
    if start_error > tolerance or end_error > tolerance:
      return False, (
          f"Inverted Bridle branch {branch_index} does not meet its physical "
          f"endpoint (start error {start_error:.3f} m, end error {end_error:.3f} m)."
      )
  return True, ""


def prepare_selected_inverted_bridle(
    selected,
    heading_deg,
    turbine_xy,
    raster_path=None,
    selected_ground=None,
    raster_sampler=None,
):
  """Prepare a selected profile's inverted bridle for 2D/3D rendering.

  This is the single rendering entry point used by Catenary, Semi-Taut and
  Taut.  It intentionally does not alter any solver result.
  """
  if not bool(selected.get("inverted_bridle", False)):
    return None

  try:
    geometry = calculate_inverted_bridle_geometry(
        selected,
        heading_deg,
        turbine_xy,
        raster_path=raster_path,
        raster_sampler=raster_sampler,
    )

    # Catenary/Semi-Taut need the raster-following original profile for their
    # grounded DDP->TDP branch section.
    geometry["profiles"] = build_inverted_bridle_branch_profiles(
        selected,
        geometry,
        ground_profile=selected_ground,
    )

    if (
        raster_path
        and str(selected.get("system_type", "")) in ("Catenary", "Semi-Taut")
    ):
      geometry = apply_inverted_bridle_raster_profiles(
          selected,
          selected,
          geometry,
          raster_path,
          turbine_xy,
          heading_deg,
          raster_sampler=raster_sampler,
      )

    profiles_valid, profile_message = validate_inverted_bridle_branch_profiles(
        geometry
    )
    if not profiles_valid:
      raise ValueError(profile_message)

    # A branch list is mandatory when the bridle is enabled.  This catches
    # geometry failures immediately instead of silently rendering only the
    # original single mooring line.
    if len(geometry.get("profiles", [])) != 2:
      raise ValueError(
          "Inverted Bridle geometry did not produce two branch profiles."
      )

    return geometry

  except Exception as exc:
    selected["bridle_geometry_error"] = str(exc)
    return None


def apply_inverted_bridle_raster_profiles(
    data, local_result, bridle_geometry, raster_path, turbine_xy, heading_deg,
    raster_sampler=None,
):
  """Apply raster seabed interaction to each inverted-bridle branch.

  The branch elevation/profile is inherited from the solved mooring profile.
  Only the portion that is physically grounded (DDP->TDP) is replaced by the
  actual raster seabed, matching terrain_following_line_geometry().  This is
  a rendering/export transformation and does not alter the original solver.
  """
  if not bridle_geometry or not bridle_geometry.get("enabled"):
    return bridle_geometry
  if str(local_result.get("system_type", data.get("system_type", "Catenary"))) == "Taut":
    return bridle_geometry
  if not raster_path:
    return bridle_geometry

  system = str(local_result.get("system_type", data.get("system_type", "Catenary")))
  try:
    clearance = max(0.0, float(data.get("mooring_seabed_clearance", 0.02)))
  except Exception:
    clearance = 0.02

  x_ddp = None
  x_td = local_result.get("X_td")
  sub_x = np.asarray(local_result.get("sub_x", []), dtype=float)
  if system != "Taut" and sub_x.size:
    x_ddp = float(sub_x[-1])

  for branch in bridle_geometry.get("profiles", []):
    bx = np.asarray(branch.get("x", []), dtype=float)
    by = np.asarray(branch.get("y", []), dtype=float)
    bz = np.asarray(branch.get("z", []), dtype=float).copy()
    px = np.asarray(branch.get("profile_x", []), dtype=float)

    if bx.size < 2 or by.size != bx.size or bz.size != bx.size or px.size != bx.size:
      continue

    terrain_z = np.asarray(
        raster_sampler.sample(bx, by)
        if raster_sampler is not None
        else sample_raster_seabed(raster_path, bx, by),
        dtype=float,
    )
    valid = np.isfinite(terrain_z)

    if system == "Catenary":
      # Catenary has grounded chain from DDP to TDP.
      if x_ddp is not None and x_td is not None:
        ground = valid & (px >= x_ddp - 1e-9) & (px <= float(x_td) + 1e-9)
      else:
        ground = np.zeros(px.shape, dtype=bool)
    else:
      # Semi-Taut: only the bottom-chain DDP->TDP section lies on the seabed.
      if x_ddp is not None and x_td is not None:
        ground = valid & (px >= x_ddp - 1e-9) & (px <= float(x_td) + 1e-9)
      else:
        ground = np.zeros(px.shape, dtype=bool)

    bz[ground] = terrain_z[ground] + clearance

    # Smooth the transition from raster ground into the suspended branch,
    # without pulling the synthetic rope / upper section onto the seabed.
    if np.any(ground) and x_td is not None:
      span = max(abs(float(x_td) - float(x_ddp if x_ddp is not None else x_td)), 1.0)
      blend = min(max(5.0, 0.10 * span), 20.0)
      trans = (
          valid
          & (px >= float(x_td))
          & (px <= float(x_td) + blend)
      )
      if np.any(trans):
        # Recover the original solved branch elevation at the transition.
        original_z = np.asarray(branch.get("z", bz), dtype=float)
        q = np.clip((px[trans] - float(x_td)) / max(blend, 1e-9), 0.0, 1.0)
        smooth = q * q * (3.0 - 2.0 * q)
        ground_z = terrain_z[trans] + clearance
        bz[trans] = (1.0 - smooth) * ground_z + smooth * original_z[trans]

    branch["z"] = bz
    branch["terrain_z"] = terrain_z
    branch["ground_mask"] = ground

  return bridle_geometry


def calculate_inverted_bridle_loads(data, bridle_geometry):
  """Calculate a symmetric two-branch static load split at the lower joint."""
  if not bridle_geometry:
    return None
  system_type = str(data.get("system_type", "Catenary"))
  H = float(data.get("HF", data.get("HA", 0.0)))
  V_joint = np.nan
  try:
    if system_type == "Semi-Taut":
      checks = data.get("segment_tension_checks") or []
      if checks:
        V_joint = float(checks[0].get("V_end", np.nan))
    else:
      x_plot = np.asarray(data.get("x_plot", []), dtype=float)
      z_plot = np.asarray(data.get("z_plot", []), dtype=float)
      sub_x = np.asarray(data.get("sub_x", []), dtype=float)
      sub_z = np.asarray(data.get("sub_z", []), dtype=float)
      if sub_x.size > 0 and sub_z.size == sub_x.size:
        full_x = np.concatenate((sub_x, x_plot[1:]))
        full_z = np.concatenate((sub_z, z_plot[1:]))
      else:
        full_x, full_z = x_plot, z_plot
      if full_x.size >= 2:
        idx = int(np.argmin(np.hypot(
            full_x - bridle_geometry["joint_local"][0],
            full_z - bridle_geometry["joint_local"][1],
        )))
        s_from_anchor = _profile_arc_length(full_x, full_z)
        s_from_fairlead = max(s_from_anchor[-1] - s_from_anchor[idx], 0.0)
        d = float(data.get("chain_d", data.get("rope_d", 0.0)))
        if system_type == "Catenary":
          W = 0.75 * np.pi * (d / 2.0) ** 2 * (7850.0 - 1025.0) * 9.81
        else:
          W = np.pi * (d / 2.0) ** 2 * (970.0 - 1025.0) * 9.81
        V_joint = float(data.get("VF", 0.0)) - W * s_from_fairlead
  except Exception:
    V_joint = np.nan
  if not np.isfinite(V_joint):
    V_joint = float(data.get("VF", 0.0))
  trunk_tension = float(np.hypot(H, V_joint))
  theta = np.radians(float(bridle_geometry["inter_anchor_angle_deg"]))
  branch_tension = trunk_tension / max(2.0 * np.cos(theta / 2.0), 1e-9)
  branches = []
  jx, jy, jz = bridle_geometry["joint"]
  for idx, a in enumerate(bridle_geometry["anchors"], start=1):
    vec = np.array([a["x"] - jx, a["y"] - jy, a["z"] - jz], dtype=float)
    L = float(np.linalg.norm(vec))
    if L <= 1e-9:
      continue
    tv = branch_tension * vec / L
    branches.append({"branch": idx, "tension": branch_tension, "length": L,
                     "HA": float(np.hypot(tv[0], tv[1])), "VA": float(tv[2])})
  return {"joint_tension": trunk_tension, "joint_H": H, "joint_V": float(V_joint),
          "branch_tension": branch_tension, "branches": branches}


def get_cell_val(sheet, cell_address, default=0.0, cast_type=float):
  """Safely reads an Excel cell, cleaning formatting (commas/currency) and casting."""
  try:
    val = sheet.range(cell_address).value
    if val is None or val == "":
      return default
    if isinstance(val, str):
      cleaned = (
          val.replace(",", "")
          .replace("$", "")
          .replace("€", "")
          .replace("£", "")
          .strip()
      )
      if not cleaned:
        return default
      return cast_type(cleaned)
    return cast_type(val)
  except Exception as e:
    print(
        f"[WARNING] Cell {cell_address} read error: {e}. Falling back to"
        f" default: {default}"
    )
    return default


MIN_LINE_MBL_FOS = 1.50  # preliminary screening factor; not a code design factor

# Semi-Taut geometry protection settings.
# The first value reserves a physical length of BOTTOM CHAIN in suspension
# between the TDP and the lower chain/rope joint.  The second value prevents
# the solved synthetic-rope section from being allowed to touch the nominal
# seabed in the static design profile.
SEMI_TAUT_MIN_SUSPENDED_BOTTOM_CHAIN = 10.0  # m
SEMI_TAUT_MIN_BOTTOM_CHAIN_FRACTION = 0.05  # fraction of bottom-chain length
SEMI_TAUT_MIN_ROPE_SEABED_CLEARANCE = 2.0  # m
SEMI_TAUT_MIN_TAUT_TENSION = 100.0  # N
TAUT_MIN_TENSION = 100.0  # N
TENSION_PROFILE_POINTS = 201
GEOMETRY_ENDPOINT_TOLERANCE = 0.50  # m
GROUND_LINE_MIN_RATIO = 0.15  # fail when less than 15% of the net/main line is grounded
GROUND_LINE_MAX_RATIO = 0.85  # high-ground-line failure threshold
GROUND_LINE_VERTICAL_ANGLE_DEG = 80.0  # fail high-ground-line case only when suspended line is near-vertical

def validate_mooring_geometry(x_plot, z_plot, xf, zf, expected_start=(0.0, 0.0), label="Mooring"):
  """Validate that a solved profile is finite, continuous and reaches both endpoints.

  This deliberately validates the *calculated* profile; it must not be repaired by
  forcing the final point onto the fairlead because that can hide a broken solution.
  """
  x = np.asarray(x_plot, dtype=float)
  z = np.asarray(z_plot, dtype=float)
  if x.size < 3 or z.size != x.size:
    return False, f"{label} geometry is incomplete."
  if not np.all(np.isfinite(x)) or not np.all(np.isfinite(z)):
    return False, f"{label} geometry contains non-finite coordinates."
  ds = np.hypot(np.diff(x), np.diff(z))
  if np.any(ds <= 1e-9):
    return False, f"{label} geometry contains zero-length segments."
  start_err = float(np.hypot(x[0]-expected_start[0], z[0]-expected_start[1]))
  end_err = float(np.hypot(x[-1]-xf, z[-1]-zf))
  if start_err > GEOMETRY_ENDPOINT_TOLERANCE:
    return False, f"{label} does not start at the padeye connection (error={start_err:.2f} m)."
  if end_err > GEOMETRY_ENDPOINT_TOLERANCE:
    return False, f"{label} does not connect to the fairlead (error={end_err:.2f} m)."
  if np.any(np.diff(x) < -1e-6):
    return False, f"{label} reverses horizontally, indicating a broken/kinked solution."
  return True, ""

def calculate_suspended_line_angle_deg(water_depth, fairlead_draft, fairlead_x, tdp_x):
  """Angle of the suspended line from horizontal, using water depth.

  The vertical rise is measured from the seabed/TDP elevation (-water_depth)
  to the fairlead elevation (fairlead_draft). The horizontal run is from the
  TDP to the fairlead. This is the same angle used for the near-vertical
  ground-line feasibility check.
  """
  horizontal_run = abs(float(fairlead_x) - float(tdp_x))
  vertical_rise = abs(float(fairlead_draft) - (-abs(float(water_depth))))
  return float(np.degrees(np.arctan2(vertical_rise, max(horizontal_run, 1e-12))))


def read_padeye_parameters(sheet):
  """Read Catenary/Semi-Taut padeye inputs from the specified Excel cells."""
  params = {
      "DEA": {
          "angle_deg": excel_cell(sheet, "dea_padeye_angle", default=15.0, cast_type=float),
      },
      "Suction Pile": {
          "angle_deg": excel_cell(sheet, "suction_padeye_angle", default=15.0, cast_type=float),
          "position_fraction": excel_cell(sheet, "suction_padeye_position", default=0.50, cast_type=float),
      },
      "Driven Pile": {
          "angle_deg": excel_cell(sheet, "driven_padeye_angle", default=15.0, cast_type=float),
          "position_fraction": excel_cell(sheet, "driven_padeye_position", default=0.50, cast_type=float),
      },
      "Drilled Pile": {
          "angle_deg": excel_cell(sheet, "drilled_padeye_angle", default=15.0, cast_type=float),
          "position_fraction": excel_cell(sheet, "drilled_padeye_position", default=0.50, cast_type=float),
      },
  }
  params = _validate_padeye_parameters(params)
  print("[INFO] Padeye parameters read from Excel:")
  for name, cfg in params.items():
    pos = cfg.get("position_fraction")
    pos_txt = f", position={100.0*pos:.1f}% down from top" if pos is not None else ""
    print(f"  - {name}: angle={cfg['angle_deg']:.2f} deg from horizontal{pos_txt}")
  return params


def safe_list_val(item, default=0.0, cast_type=float):
  """Safely converts an item from an Excel range list, handling None/strings."""
  if item is None or item == "":
    return default
  try:
    if isinstance(item, str):
      cleaned = (
          item.replace(",", "")
          .replace("$", "")
          .replace("€", "")
          .replace("£", "")
          .strip()
      )
      if not cleaned:
        return default
      return cast_type(cleaned)
    return cast_type(item)
  except:
    return default


def read_anchor_database(sheet):
  try:
    areas = excel_range(sheet, "anchor_areas")
    depths = excel_range(sheet, "anchor_depths")
    heights = excel_range(sheet, "anchor_heights")
    widths = excel_range(sheet, "anchor_widths")

    if not isinstance(areas, list):
      areas = [areas]
    if not isinstance(depths, list):
      depths = [depths]
    if not isinstance(heights, list):
      heights = [heights]
    if not isinstance(widths, list):
      widths = [widths]

    anchor_db = {}
    all_anchors = [
        "DEA",
        "Suction Pile",
        "Driven Pile",
        "Drilled Pile",
        "Gravity",
    ]

    defaults = [
        {"area": 15.0, "depth": 8.0, "height": 3.0, "width": 4.0},
        {"area": 20.0, "depth": 15.0, "height": 10.0, "width": 5.0},
        {"area": 10.0, "depth": 20.0, "height": 15.0, "width": 2.0},
        {"area": 8.0, "depth": 25.0, "height": 18.0, "width": 2.0},
        {"area": 50.0, "depth": 5.0, "height": 5.0, "width": 10.0},
    ]

    for i, a_type in enumerate(all_anchors):
      def_vals = defaults[i]
      anchor_db[a_type] = {
          "area": safe_list_val(
              areas[i] if i < len(areas) else None, def_vals["area"]
          ),
          "depth": safe_list_val(
              depths[i] if i < len(depths) else None, def_vals["depth"]
          ),
          "height": safe_list_val(
              heights[i] if i < len(heights) else None, def_vals["height"]
          ),
          "width": safe_list_val(
              widths[i] if i < len(widths) else None, def_vals["width"]
          ),
      }
    print("[INFO] Anchor database successfully read from Excel.")
    return anchor_db, all_anchors
  except Exception as e:
    print(
        f"[WARNING] Could not read anchor database from Excel ({e}). Using"
        " defaults."
    )
    return {
        "DEA": {"area": 15.0, "depth": 8.0, "height": 3.0, "width": 4.0},
        "Suction Pile": {
            "area": 20.0,
            "depth": 15.0,
            "height": 10.0,
            "width": 5.0,
        },
        "Driven Pile": {
            "area": 10.0,
            "depth": 20.0,
            "height": 15.0,
            "width": 2.0,
        },
        "Drilled Pile": {
            "area": 8.0,
            "depth": 25.0,
            "height": 18.0,
            "width": 2.0,
        },
        "Gravity": {"area": 50.0, "depth": 5.0, "height": 5.0, "width": 10.0},
    }, ["DEA", "Suction Pile", "Driven Pile", "Drilled Pile", "Gravity"]


def get_anchor_dimensions(anchor_db, primary_anc):
  dims = anchor_db.get(
      primary_anc, {"depth": 8.0, "height": 3.0, "width": 4.0}
  )
  return (
      dims.get("width", 2.0),
      dims.get("height", 2.0),
      dims.get("depth", 8.0),
  )


def evaluate_anchor_capacities(anchor_db, active_anchors, HA, VA):
  if not anchor_db:
    return {}
  Su, Nc, alpha, FoS_req = 50000.0, 9.0, 0.8, 1.5
  results = {}

  for a_type in active_anchors:
    if a_type not in anchor_db:
      continue
    dims = anchor_db[a_type]
    area, depth = float(dims["area"]), float(dims["depth"])

    fos, status, msg = 0.0, "FAIL", ""
    if area <= 0.0 or depth <= 0.0:
      results[a_type] = {
          "status": "FAIL", "fos": 0.0, "msg": "Anchor capacity input area/depth must be > 0.",
          "area": area, "depth": depth,
      }
      continue

    if a_type == "DEA":
      if VA > 500.0:
        msg = f"Uplift limit exceeded (VA = {VA/1000:.1f} kN)"
      else:
        uhc = area * Nc * Su
        fos = uhc / abs(HA) if HA != 0 else float("inf")
        status, msg = (
            ("PASS" if fos >= FoS_req else "FAIL"),
            f"UHC: {uhc/1000:.1f} kN | FoS: {fos:.2f}",
        )
    elif a_type in ["Suction Pile", "Driven Pile", "Drilled Pile"]:
      D = math.sqrt(4.0 * area / math.pi)
      uhc_v = (math.pi * D * depth * alpha * Su) + (area * Nc * Su)
      uhc_h = 9.0 * Su * D * depth
      unity = (abs(VA) / uhc_v) ** 2 + (abs(HA) / uhc_h) ** 2
      fos = (1.0 / math.sqrt(unity)) if unity > 0 else float("inf")
      status, msg = (
          ("PASS" if fos >= FoS_req else "FAIL"),
          f"Combined P-FoS: {fos:.2f} (H FoS={uhc_h/max(abs(HA),1e-9):.2f}, V FoS={uhc_v/max(abs(VA),1e-9):.2f})",
      )
    elif a_type == "Gravity":
      submerged_weight = area * depth * (2400 - 1025) * 9.81
      if abs(VA) > submerged_weight:
        msg = "Uplift exceeds weight"
      else:
        fos = (submerged_weight * 0.5) / abs(HA) if HA != 0 else float("inf")
        status, msg = (
            ("PASS" if fos >= FoS_req else "FAIL"),
            f"Sliding FoS: {fos:.2f}",
        )

    results[a_type] = {
        "status": status,
        "fos": fos,
        "msg": msg,
        "area": area,
        "depth": depth,
        "HA_N": float(HA),
        "VA_N": float(VA),
    }
  return results


def _workbook_path_matches_selected_path(workbook, selected_path):
  """Return whether Excel reports the selected workbook's exact file path."""
  try:
    return _normalise_excel_workbook_path(workbook.fullname) == selected_path
  except Exception:
    return False


def _workbook_name_matches_selected_path(workbook, selected_path):
  """Compare workbook names for the macOS OneDrive/iCloud fallback."""
  try:
    selected_name = unicodedata.normalize(
        "NFC", os.path.basename(selected_path)
    ).casefold()
    workbook_name = unicodedata.normalize(
        "NFC", str(workbook.name)
    ).casefold()
    return workbook_name == selected_name
  except Exception:
    return False


def _find_open_selected_excel_workbook(selected_path, sheet_name):
  """Return the selected open book without reopening it in Excel.

  Windows retains strict full-path matching.  macOS first tries that same
  match, then supports Excel's cloud/HFS path representation by matching the
  filename and immediately confirming that the required worksheet exists.
  """
  candidates = []

  # Prefer the workbook active in Excel on macOS when it matches by name.  It
  # avoids reading the fragile ``fullname`` property for a cloud workbook.
  if sys.platform == "darwin":
    try:
      active_book = xw.books.active
      if _workbook_name_matches_selected_path(active_book, selected_path):
        try:
          active_book.sheets[sheet_name]
          return active_book
        except Exception:
          pass
      candidates.append(active_book)
    except Exception:
      pass

  try:
    for app in xw.apps:
      candidates.extend(list(app.books))
  except Exception:
    pass

  # xlwings can sometimes expose only the active collection even when app
  # enumeration is unavailable, particularly with Excel for macOS.
  try:
    active_book = xw.books.active
    candidates.append(active_book)
  except Exception:
    pass

  # Preserve the existing exact-path behavior first on every platform.
  for candidate in candidates:
    if _workbook_path_matches_selected_path(candidate, selected_path):
      return candidate

  if sys.platform == "darwin":
    # Excel for macOS can expose OneDrive/iCloud books with a different HFS,
    # URL or internal path.  Do not read ``fullname`` in this branch: name +
    # verified worksheet is the reliable identity available to xlwings.
    for candidate in candidates:
      if not _workbook_name_matches_selected_path(candidate, selected_path):
        continue
      try:
        candidate.sheets[sheet_name]
        return candidate
      except Exception:
        continue
  return None


class _OpenpyxlRangeAdapter:
  """Expose the small ``xlwings.Range`` interface used by this application."""
  def __init__(self, worksheet, address):
    self._worksheet = worksheet
    self._address = address

  @property
  def value(self):
    cells = self._worksheet[self._address]
    if not isinstance(cells, tuple):
      return cells.value

    # openpyxl returns a tuple of tuples for a multi-cell range.  Match
    # xlwings' convenient scalar-list shape for one-row/one-column ranges,
    # while retaining a 2D list for a genuinely rectangular selection.
    if cells and isinstance(cells[0], tuple):
      values = [[cell.value for cell in row] for row in cells]
      if len(values) == 1:
        return values[0]
      if all(len(row) == 1 for row in values):
        return [row[0] for row in values]
      return values
    return [cell.value for cell in cells]


class _OpenpyxlSheetAdapter:
  """Read-only worksheet adapter for the existing ``sheet.range(...).value`` calls."""
  def __init__(self, workbook, worksheet):
    # Keep the workbook alive for the full dashboard session.
    self._workbook = workbook
    self._worksheet = worksheet

  def range(self, address):
    return _OpenpyxlRangeAdapter(self._worksheet, address)


def _open_saved_excel_sheet_without_automation(excel_path, sheet_name):
  """Read a saved .xlsx/.xlsm workbook without needing Excel automation."""
  extension = os.path.splitext(excel_path)[1].lower()
  if extension not in {".xlsx", ".xlsm"}:
    raise ValueError(
        "Direct macOS workbook reading supports .xlsx and .xlsm files only."
    )
  try:
    from openpyxl import load_workbook
  except ImportError as exc:
    raise RuntimeError(
        "openpyxl is required for the macOS no-automation fallback. "
        "Install it with: "
        f'"{sys.executable}" -m pip install openpyxl'
    ) from exc

  workbook = load_workbook(
      excel_path,
      data_only=True,
      # This application only reads cell values.  Streaming avoids loading a
      # large workbook into memory and keeps macOS independent of Excel.
      read_only=True,
      keep_links=False,
  )
  if sheet_name not in workbook.sheetnames:
    workbook.close()
    raise KeyError(f"Worksheet '{sheet_name}' was not found in the workbook.")
  return _OpenpyxlSheetAdapter(workbook, workbook[sheet_name])


def get_excel_sheet(sheet_name="Mooring_Python", excel_path=None):
  """Return ``sheet_name`` from the workbook selected on the welcome screen.

  An explicit/user-selected path deliberately takes priority over Excel's
  active workbook.  Without that ordering, opening another workbook while the
  dashboard is running can silently change the design data being analysed.
  The old active/caller/default sequence remains available for legacy callers
  which do not yet supply a selected workbook.
  """
  selected_path = excel_path or _SELECTED_EXCEL_WORKBOOK_PATH

  if selected_path:
    selected_path = _normalise_excel_workbook_path(selected_path)
    if not os.path.isfile(selected_path):
      raise FileNotFoundError(f"Excel workbook was not found: {selected_path}")

    # Excel automation on macOS requires an Apple-events permission prompt
    # which may be absent or blocked by an organisation's device policy.  The
    # dashboard only needs saved worksheet values, so prefer a direct reader
    # there.  Excel/xlwings remains an optional fallback for legacy formats or
    # unusual workbooks that openpyxl cannot read.  Windows deliberately keeps
    # its existing live-Excel behaviour, including unsaved workbook edits.
    mac_direct_read_error = None
    if sys.platform == "darwin":
      try:
        sheet = _open_saved_excel_sheet_without_automation(
            selected_path, sheet_name
        )
        print(
            "[INFO] Reading the saved selected workbook directly on macOS; "
            "Excel automation is not required."
        )
        return sheet
      except Exception as direct_read_exc:
        mac_direct_read_error = direct_read_exc

    # Reuse a copy already open in Excel.  This preserves unsaved edits and
    # supports macOS cloud-workbook paths that do not string-match Finder's
    # POSIX selection path.
    workbook = _find_open_selected_excel_workbook(selected_path, sheet_name)

    try:
      if workbook is None:
        print(f"[INFO] Opening selected Excel workbook: {selected_path}")
        if sys.platform == "darwin":
          # ``Books.open`` is the xlwings API that returns an already-open
          # book when possible and is more reliable than ``xw.Book(fullpath)``
          # with Excel for macOS.
          try:
            workbook = xw.apps.active.books.open(selected_path)
          except Exception:
            workbook = xw.books.open(selected_path)
        else:
          workbook = xw.Book(selected_path)
      sheet = workbook.sheets[sheet_name]
      print(
          "[INFO] Successfully connected to selected workbook sheet:"
          f" '{sheet_name}'"
      )
      return sheet
    except Exception as exc:
      fallback_error = mac_direct_read_error

      details = str(exc).lower()
      if sys.platform == "darwin" and (
          "-1743" in details
          or "not authorized" in details
          or "not permitted to send apple events" in details
      ):
        raise PermissionError(
            "macOS has blocked the Python application from controlling "
            "Microsoft Excel. Allow it under System Settings > Privacy & "
            "Security > Automation, then restart the application."
        ) from exc
      raise ValueError(
          f"Workbook '{selected_path}' does not contain an accessible "
          f"'{sheet_name}' worksheet ({exc})"
          + (
              f". Direct saved-workbook reading also failed ({fallback_error})."
              if fallback_error is not None else "."
          )
      ) from exc

  try:
    wb = xw.books.active
    if wb:
      sheet = wb.sheets[sheet_name]
      print(
          "[INFO] Successfully connected to active workbook sheet:"
          f" '{sheet_name}'"
      )
      return sheet
  except Exception:
    pass

  try:
    sheet = xw.Book.caller().sheets[sheet_name]
    print(
        "[INFO] Successfully connected via xw.Book.caller() to sheet:"
        f" '{sheet_name}'"
    )
    return sheet
  except Exception:
    pass

  try:
    print(f"[INFO] Opening workbook directly from path: {DEFAULT_EXCEL_WORKBOOK_PATH}")
    wb = xw.Book(DEFAULT_EXCEL_WORKBOOK_PATH)
    sheet = wb.sheets[sheet_name]
    return sheet
  except Exception as exc:
    print(f"[ERROR] Could not open workbook at {DEFAULT_EXCEL_WORKBOOK_PATH}: {exc}")
    raise


# =========================================================================
# TRUE MULTI-SEGMENT CATENARY SOLVER (SEMI-TAUT)
# =========================================================================
def eval_multisegment_profile(H, U, segments, num_points=50):
  L_ground = U if U > 0 else 0.0
  V_anchor = -U if U <= 0 else 0.0

  x_vals, z_vals, joints = [], [], [(0.0, 0.0)]
  X_tot, Z_tot, V_curr = 0.0, 0.0, 0.0

  for i, seg in enumerate(segments):
    L_seg, w, EA = seg["L"], seg["w"], seg["EA"]

    if i == 0 and L_ground > 0:
      s_g = np.linspace(0, L_ground, num_points)
      x_g = X_tot + s_g * (1.0 + H / EA)
      x_vals.extend(x_g)
      z_vals.extend(np.full_like(s_g, Z_tot))
      X_tot = x_g[-1]
      V_curr = 0.0
      L_sus = L_seg - L_ground
    else:
      L_sus = L_seg
      if i == 0:
        V_curr = V_anchor

    if L_sus > 0:
      s_s = np.linspace(0, L_sus, num_points)
      if len(x_vals) > 0:
        s_s = s_s[1:]

      if len(s_s) > 0:
        V_start = V_curr
        V_arr = V_start + w * s_s

        T_arr = np.sqrt(H**2 + V_arr**2)
        T_start = np.sqrt(H**2 + V_start**2)

        dx = (H / w) * (
            np.arcsinh(V_arr / H) - np.arcsinh(V_start / H)
        ) + (H * s_s) / EA
        dz = (T_arr - T_start) / w + (V_start * s_s + 0.5 * w * s_s**2) / EA

        x_vals.extend(X_tot + dx)
        z_vals.extend(Z_tot + dz)

        X_tot += dx[-1]
        Z_tot += dz[-1]
        V_curr = V_start + w * L_sus

    joints.append((X_tot, Z_tot))

  return np.array(x_vals), np.array(z_vals), V_anchor, V_curr, L_ground, joints


# =========================================================================
# EVALUATION FUNCTIONS FOR UNIFIED DASHBOARD
# =========================================================================
# Preliminary static tension screening for Semi-Taut segment MBL checks.
def calculate_segment_tension_checks(H, U, segments):
  """Return start/end/max tension for each Semi-Taut segment."""
  L_ground = max(float(U), 0.0)
  V_curr = -float(U) if U <= 0.0 else 0.0
  checks = []
  for i, seg in enumerate(segments):
    L_seg = float(seg["L"])
    w = float(seg["w"])
    if L_seg <= 0.0:
      checks.append({"segment": i, "L": L_seg, "T_start": np.nan, "T_end": np.nan, "T_max": np.nan})
      continue
    L_sus = L_seg - L_ground if i == 0 else L_seg
    L_sus = max(L_sus, 0.0)
    if i == 0 and L_ground > 0.0:
      V_start = 0.0
    else:
      V_start = V_curr
    V_end = V_start + w * L_sus
    T_ground = abs(H)
    T_start = max(T_ground, float(np.hypot(H, V_start)))
    T_end = max(T_ground, float(np.hypot(H, V_end)))
    checks.append({
        "segment": i, "L": L_seg, "L_sus": L_sus,
        "T_start": T_start, "T_end": T_end, "T_max": max(T_start, T_end),
        "V_start": V_start, "V_end": V_end,
    })
    V_curr = V_end
    L_ground = 0.0
  return checks


def evaluate_catenary_anchor(
    water_depth,
    anchor_radius,
    line_length,
    chain_d,
    fairlead_radius,
    fairlead_draft,
    num_turbines,
    farm_area,
    num_lines,
    center_lat,
    center_lon,
    buffer_zone,
    anchor_db,
    primary_anc,
    padeye_params=None,
):
  a_width, a_height, a_depth = get_anchor_dimensions(anchor_db, primary_anc)

  # Padeye placed on the inside of the anchor
  r_padeye = anchor_radius - a_width
  xf, zf = abs(r_padeye - fairlead_radius), fairlead_draft - (-water_depth)

  sub_x, sub_z, L_sub, x_sub_span, z_anchor, padeye_angle_deg = get_subsurface_geometry(
      water_depth, anchor_db, primary_anc, padeye_params=padeye_params
  )
  L_main, xf_net = line_length - L_sub, xf - x_sub_span
  min_geometric_main = float(np.hypot(xf_net, zf))
  min_geometric_length = float(L_sub + min_geometric_main)
  excess_line = float(line_length - min_geometric_length)

  if xf_net <= 0.0:
    return {"success": False, "status": "FAIL", "msg": "Invalid net suspended span after padeye/subsurface geometry."}
  if line_length <= 0.0 or L_main <= 0.0 or line_length <= min_geometric_length:
    return {
        "success": False,
        "status": "FAIL",
        "msg": (f"Line length too short: actual={line_length:.2f} m, "
                 f"minimum geometric length={min_geometric_length:.2f} m."),
    }

  metal_area = 0.75 * np.pi * (chain_d / 2.0) ** 2
  W, EA = metal_area * (7850.0 - 1025.0) * 9.81, metal_area * 210e9

  try:
    HA, VA, HF, VF, info = mc.catenary(
        XF=xf_net, ZF=zf, L=L_main, EA=EA, W=W, depth=water_depth
    )
    # Always calculate and retain the actual H/V/resultant tensions before
    # applying any failure criterion, so failed configurations can still be
    # fully diagnosed in the Failure Criteria report.
    T_fairlead = float(np.hypot(HF, VF))
    T_padeye = float(np.hypot(HA, VA))
    T_max = max(T_fairlead, T_padeye)
    if VA > 0.0 and primary_anc == "DEA":
      chain_grade_diag, chain_mbl_diag, _ = select_optimal_chain_grade(
          chain_d, T_max, min_fos=MIN_LINE_MBL_FOS
      )
      return {
          "success": False,
          "status": "FAIL",
          "msg": f"Uplift VA = {VA:.2f} N (> 0).",
          "HA": HA, "VA": VA, "HF": HF, "VF": VF,
          "T_padeye": T_padeye, "T_fairlead": T_fairlead, "T_max": T_max,
          "line_fos_padeye": chain_mbl_diag / max(T_padeye, 1e-9),
          "line_fos_fairlead": chain_mbl_diag / max(T_fairlead, 1e-9),
          "chain_grade": chain_grade_diag,
      }
    chain_grade, chain_mbl_N, _ = select_optimal_chain_grade(
        chain_d, T_max, min_fos=MIN_LINE_MBL_FOS
    )
    chain_fos_fairlead = chain_mbl_N / max(T_fairlead, 1e-9)
    chain_fos_padeye = chain_mbl_N / max(T_padeye, 1e-9)
    if chain_fos_fairlead < MIN_LINE_MBL_FOS or chain_fos_padeye < MIN_LINE_MBL_FOS:
      return {
          "success": False,
          "status": "FAIL",
          "msg": (f"Chain MBL screening failed: fairlead FoS={chain_fos_fairlead:.2f}, "
                   f"padeye FoS={chain_fos_padeye:.2f} (< {MIN_LINE_MBL_FOS:.2f})."),
          "HA": HA, "VA": VA, "HF": HF, "VF": VF,
          "T_padeye": T_padeye, "T_fairlead": T_fairlead, "T_max": T_max,
          "line_fos_fairlead": chain_fos_fairlead, "line_fos_padeye": chain_fos_padeye,
          "chain_grade": chain_grade, "chain_d": chain_d, "system_type": "Catenary",
          "anchor_results": evaluate_anchor_capacities(anchor_db, [primary_anc], HA, VA),
          "min_geometric_length": min_geometric_length, "excess_line": excess_line,
      }

    anchor_results = evaluate_anchor_capacities(
        anchor_db, [primary_anc], HA, VA
    )
    if anchor_results.get(primary_anc, {}).get("status") == "FAIL":
      return {
          "success": False,
          "status": "FAIL",
          "msg": "Geotechnical capacity failed.",
          "HA": HA, "VA": VA, "HF": HF, "VF": VF,
          "T_padeye": T_padeye, "T_fairlead": T_fairlead, "T_max": T_max,
          "line_fos_fairlead": chain_fos_fairlead, "line_fos_padeye": chain_fos_padeye,
          "chain_grade": chain_grade, "chain_d": chain_d, "system_type": "Catenary",
          "anchor_results": anchor_results,
          "min_geometric_length": min_geometric_length, "excess_line": excess_line,
      }

    if x_sub_span <= 0.0 or xf_net <= 0.0:
      return {"success": False, "status": "FAIL", "msg": "Invalid net suspended span after padeye/subsurface geometry."}

    L_sus = abs(VF) / W if ("Line_U" not in info and VA <= 1.0) else L_main
    L_sus = float(np.clip(L_sus, 0.0, L_main))
    L_bot, H, a = max(L_main - L_sus, 0.0), max(abs(HF), 1.0), max(abs(HF), 1.0) / W
    ground_ratio = float(L_bot / max(L_main, 1e-12))
    X_td_net = xf_net - a * np.arcsinh(L_sus / a)
    suspended_line_angle_deg = calculate_suspended_line_angle_deg(
        water_depth, fairlead_draft, xf_net, X_td_net
    )
    ground_line_low_fail = ground_ratio < GROUND_LINE_MIN_RATIO
    ground_line_high_fail = (
        ground_ratio > GROUND_LINE_MAX_RATIO
        and suspended_line_angle_deg >= GROUND_LINE_VERTICAL_ANGLE_DEG
    )
    ground_line_fail = ground_line_low_fail or ground_line_high_fail
    ground_line_warning = False
    ground_line_warning_text = ""
    if ground_line_low_fail:
      ground_line_fail_msg = (
          f"Insufficient ground line: {ground_ratio*100.0:.1f}% of the net/main line is grounded "
          f"(minimum={GROUND_LINE_MIN_RATIO*100.0:.0f}%)."
      )
    elif ground_line_high_fail:
      ground_line_fail_msg = (
          f"Excessive ground line with near-vertical suspended section: "
          f"{ground_ratio*100.0:.1f}% grounded and suspended-line angle="
          f"{suspended_line_angle_deg:.1f}° (vertical threshold={GROUND_LINE_VERTICAL_ANGLE_DEG:.1f}°, "
          f"ground threshold={GROUND_LINE_MAX_RATIO*100.0:.0f}%)."
      )
    else:
      ground_line_fail_msg = ""
    if ground_line_fail:
      return {
          "success": False,
          "status": "FAIL",
          "msg": ground_line_fail_msg,
          "ground_ratio": ground_ratio,
          "suspended_line_angle_deg": suspended_line_angle_deg,
          "min_geometric_length": min_geometric_length,
          "excess_line": excess_line,
          "L_bot": L_bot, "L_sus": L_sus,
          "grounded_surface_length": L_bot,
          "suspended_surface_length": L_sus,
          "HA": HA, "VA": VA, "HF": HF, "VF": VF,
          "T_padeye": T_padeye, "T_fairlead": T_fairlead, "T_max": T_max,
          "line_fos_fairlead": chain_fos_fairlead,
          "line_fos_padeye": chain_fos_padeye,
          "required_line_fos": MIN_LINE_MBL_FOS,
          "chain_grade": chain_grade,
          "chain_d": chain_d, "system_type": "Catenary",
          "anchor_results": anchor_results,
          "primary_anc": primary_anc,
      }

    # Build the suspended catenary without evaluating cosh() on the grounded
    # portion. np.where() evaluates both branches and can therefore overflow
    # even when the suspended branch is not used.
    x_plot_net = np.linspace(0.0, xf_net, 500)
    z_plot_net = np.full_like(x_plot_net, -water_depth, dtype=float)
    suspended_mask = x_plot_net > X_td_net
    if np.any(suspended_mask):
      arg = (x_plot_net[suspended_mask] - X_td_net) / a
      if not np.all(np.isfinite(arg)) or np.max(np.abs(arg)) > 700.0:
        return {"success": False, "status": "FAIL",
                "msg": "Catenary solution is numerically unstable (excessive catenary curvature)."}
      z_plot_net[suspended_mask] = -water_depth + a * (np.cosh(arg) - 1.0)

    # Do NOT force the last point to the fairlead.  A forced endpoint can make
    # a broken/short line appear connected. The calculated endpoint must agree
    # with the fairlead within the geometry tolerance.
    # x_plot_net / z_plot_net are in absolute plot coordinates:
    #   x: net horizontal distance from the subsurface/DDP reference
    #   z: absolute elevation, with seabed at -water_depth.
    # Therefore the endpoint elevation must be compared with the actual
    # fairlead elevation (fairlead_draft), NOT zf (which is the relative
    # vertical separation used by MoorPy).
    geometry_ok, geometry_msg = validate_mooring_geometry(
        x_plot_net, z_plot_net, xf_net, fairlead_draft,
        expected_start=(0.0, -water_depth), label="Catenary"
    )
    if not geometry_ok:
      return {"success": False, "status": "FAIL", "msg": geometry_msg,
              "HA": HA, "VA": VA, "HF": HF, "VF": VF,
              "T_padeye": T_padeye, "T_fairlead": T_fairlead, "T_max": T_max,
              "line_fos_fairlead": chain_fos_fairlead, "line_fos_padeye": chain_fos_padeye,
              "chain_grade": chain_grade, "chain_d": chain_d, "system_type": "Catenary",
              "anchor_results": anchor_results, "ground_ratio": ground_ratio,
              "grounded_surface_length": L_bot, "suspended_surface_length": L_sus,
              "suspended_line_angle_deg": suspended_line_angle_deg,
              "min_geometric_length": min_geometric_length, "excess_line": excess_line}

    return {
        "success": True,
        "status": "PASS",
        "x_plot": x_plot_net + x_sub_span,
        "z_plot": z_plot_net,
        "xf": xf,
        "water_depth": water_depth,
        "fairlead_draft": fairlead_draft,
        "anchor_radius": anchor_radius,
        "line_length": line_length,
        "num_lines": num_lines,
        "num_turbines": num_turbines,
        "farm_area": farm_area,
        "buffer_zone": buffer_zone,
        "L_sus": L_sus,
        "L_bot": L_bot,
        "HF": HF,
        "VF": VF,
        "HA": HA,
        "VA": VA,
        "T_fairlead": T_fairlead,
        "T_padeye": T_padeye,
        "T_max": T_max,
        "line_fos_fairlead": chain_fos_fairlead,
        "line_fos_padeye": chain_fos_padeye,
        "required_line_fos": MIN_LINE_MBL_FOS,
        "padeye_angle_deg": padeye_angle_deg,
        "padeye_z": z_anchor,
        "padeye_position_fraction": (padeye_params or {}).get(primary_anc, {}).get("position_fraction", None),
        "X_td": X_td_net + x_sub_span,
        "title": f"Catenary Analysis - {primary_anc}",
        "center_lat": center_lat,
        "center_lon": center_lon,
        "sub_x": sub_x,
        "sub_z": sub_z,
        "L_sub": L_sub,
        "min_geometric_length": min_geometric_length,
        "min_geometric_main_length": min_geometric_main,
        "excess_line": excess_line,
        "ground_ratio": ground_ratio,
        "ground_line_warning": False,
        "ground_line_warning_text": "",
        "ground_line_min_ratio": GROUND_LINE_MIN_RATIO,
        "ground_line_max_ratio": GROUND_LINE_MAX_RATIO,
        "suspended_line_angle_deg": suspended_line_angle_deg,
        "ground_line_vertical_angle_deg": GROUND_LINE_VERTICAL_ANGLE_DEG,
        "grounded_surface_length": L_bot,
        "suspended_surface_length": L_sus,
        "chain_d": chain_d,
        "system_type": "Catenary",
        "chain_grade": chain_grade,
        "anchor_results": anchor_results,
        "anchor_width": a_width,
        "anchor_height": a_height,
        "anchor_depth": a_depth,
        "primary_anc": primary_anc,
    }
  except Exception as e:
    return {"success": False, "status": "FAIL", "msg": str(e)}



def evaluate_semi_taut_anchor(
    water_depth,
    anchor_radius,
    line_length,
    rope_d,
    chain_d,
    fairlead_radius,
    fairlead_draft,
    num_turbines,
    farm_area,
    num_lines,
    center_lat,
    center_lon,
    buffer_zone,
    anchor_db,
    primary_anc,
    padeye_params=None,
    taut_percentage=30.0,
    upper_equals_lower=True,
    lower_joint_pos=0.50,
):
  """Evaluate a three-segment Semi-Taut line.

  Segment order is: bottom chain -> synthetic rope -> top chain.  H16 always
  defines the percentage of the *suspended* line occupied by the synthetic
  rope.  When H19/``upper_equals_lower`` is TRUE, the remaining suspended
  chain length is split equally between the lower and upper chain.  When H19
  is FALSE, H21/``lower_joint_pos`` fixes the Bottom Chain / Rope Joint as a
  fraction of the suspended Fairlead->TDP span (measured from the fairlead);
  the bottom-chain, rope and top-chain lengths are then derived backwards
  from that imposed joint position.
  """
  a_width, a_height, a_depth = get_anchor_dimensions(anchor_db, primary_anc)

  # Padeye placed on the inside of the anchor.
  r_padeye = anchor_radius - a_width
  xf, zf = abs(r_padeye - fairlead_radius), fairlead_draft - (-water_depth)

  sub_x, sub_z, L_sub, x_sub_span, z_anchor, padeye_angle_deg = get_subsurface_geometry(
      water_depth, anchor_db, primary_anc, padeye_params=padeye_params
  )
  line_length_main, xf_net = line_length - L_sub, xf - x_sub_span
  min_geometric_main = float(np.hypot(xf_net, zf))
  min_geometric_length = float(L_sub + min_geometric_main)
  excess_line = float(line_length - min_geometric_length)

  if xf_net <= 0.0:
    return {"success": False, "status": "FAIL",
            "msg": "Invalid net suspended span after padeye/subsurface geometry."}

  if line_length <= min_geometric_length:
    return {"success": False, "status": "FAIL",
            "msg": (f"Target span exceeds line length: actual={line_length:.2f} m, "
                     f"minimum geometric length={min_geometric_length:.2f} m.")}

  W_chain = 0.75 * np.pi * (chain_d / 2.0) ** 2 * (7850.0 - 1025.0) * 9.81
  EA_chain = 0.75 * np.pi * (chain_d / 2.0) ** 2 * 210e9
  W_rope = np.pi * (rope_d / 2.0) ** 2 * (1380.0 - 1025.0) * 9.81
  EA_rope = np.pi * (rope_d / 2.0) ** 2 * 5e9

  # H16 always controls the percentage of the SUSPENDED section occupied by
  # the taut/synthetic rope.  H19 then decides how the remaining suspended
  # chain is allocated:
  #
  #   H19=TRUE : equal chain split
  #       bottom = top = (1 - H16) / 2
  #
  #   H19=FALSE: H21 first fixes the Bottom Chain / Rope Joint position.
  #       H21 is measured from the FAIRLEAD back to the TDP over the suspended
  #       line only. Therefore:
  #         bottom chain = 1 - H21
  #         rope         = H16
  #         top chain    = H21 - H16
  #
  # The grounded portion U remains part of the physical bottom-chain length,
  # but is excluded from the H21 suspended-line fraction.
  target_taut = float(taut_percentage) / 100.0
  if not np.isfinite(target_taut) or not (0.0 < target_taut < 1.0):
    return {"success": False, "status": "FAIL",
            "msg": f"Invalid Semi-Taut taut percentage: {taut_percentage}. Expected 0-100%."}

  use_equal_chain_split = bool(upper_equals_lower)
  joint_fraction = float(lower_joint_pos)
  if not np.isfinite(joint_fraction) or not (0.0 < joint_fraction < 1.0):
    return {"success": False, "status": "FAIL",
            "msg": f"Invalid H21 lower-joint position: {lower_joint_pos}. Expected 0-1."}

  if (not use_equal_chain_split) and joint_fraction <= target_taut:
    return {"success": False, "status": "FAIL",
            "msg": (f"Invalid H19/H21/H16 combination: H21={joint_fraction:.3f} must be greater than "
                     f"the H16 rope fraction={target_taut:.3f} so a positive upper-chain section remains.")}

  # -------------------------------------------------------------------------
  # H19/H21-controlled Semi-Taut segment solver
  # -------------------------------------------------------------------------
  # H19=TRUE: retain the original equal suspended-chain split.
  # H19=FALSE: H21 controls the *geometric* position of the Lower Joint on
  # the suspended Fairlead -> TDP curve. H16 still fixes the synthetic-rope
  # fraction of the suspended line. The remaining bottom/upper chain split,
  # together with the naturally grounded length U, is solved backwards so
  # that the complete multi-segment line closes exactly and the Lower Joint
  # actually lands at H21. This removes the old situation where the segment
  # length fractions were correct but elastic stretch moved the physical joint
  # away from the requested H21 position.
  target_taut = float(taut_percentage) / 100.0
  if not np.isfinite(target_taut) or not (0.0 < target_taut < 1.0):
    return {"success": False, "status": "FAIL",
            "msg": f"Invalid Semi-Taut taut percentage: {taut_percentage}. Expected 0-100%."}

  use_equal_chain_split = bool(upper_equals_lower)
  joint_fraction = float(lower_joint_pos)
  if not np.isfinite(joint_fraction) or not (0.0 < joint_fraction < 1.0):
    return {"success": False, "status": "FAIL",
            "msg": f"Invalid H21 lower-joint position: {lower_joint_pos}. Expected 0-1."}

  if (not use_equal_chain_split) and joint_fraction <= target_taut:
    return {"success": False, "status": "FAIL",
            "msg": (f"Invalid H19/H21/H16 combination: H21={joint_fraction:.3f} must be greater than "
                     f"the H16 rope fraction={target_taut:.3f} so a positive upper-chain section remains.")}

  def make_segments(U_val):
    """Build the physical Semi-Taut segments for a given grounded length.

    H16 fixes the synthetic-rope fraction of the suspended main line.
    H19=TRUE splits the remaining chain equally.
    H19=FALSE fixes the LOWER JOINT by H21, so the suspended bottom-chain
    fraction is (1-H21), while rope=H16 and top-chain=(H21-H16).

    The grounded length U is intentionally *not* fixed by H21.  U is solved
    together with horizontal tension H so that the complete line closes to the
    fairlead.  This is what allows the ground ratio and suspended-line angle to
    change when the total line length or H16 changes.
    """
    U_pos = max(float(U_val), 0.0)
    L_suspended = line_length_main - U_pos
    if L_suspended <= 1e-8:
      return None, L_suspended

    L_rope_val = L_suspended * target_taut
    if use_equal_chain_split:
      bottom_frac = 0.5 * (1.0 - target_taut)
      top_frac = bottom_frac
    else:
      # H21 is measured from the FAIRLEAD toward the TDP.  Therefore the
      # distance from TDP to LOWER JOINT is exactly (1-H21) of the suspended
      # line, and the distance from FAIRLEAD to LOWER JOINT is H21.
      bottom_frac = 1.0 - joint_fraction
      top_frac = joint_fraction - target_taut

    if bottom_frac <= 0.0 or top_frac <= 0.0:
      return None, L_suspended

    L_bottom_suspended_val = L_suspended * bottom_frac
    L_top_chain_val = L_suspended * top_frac

    if L_bottom_suspended_val <= 0.0 or L_top_chain_val <= 0.0:
      return None, L_suspended

    L_bottom_total_val = U_pos + L_bottom_suspended_val

    segments_val = [
      {"L": L_bottom_total_val, "w": W_chain, "EA": EA_chain},
      {"L": L_rope_val, "w": W_rope, "EA": EA_rope},
      {"L": L_top_chain_val, "w": W_chain, "EA": EA_chain},
    ]
    return segments_val, L_suspended

  # Use an equivalent single-line catenary only to obtain a robust initial
  # horizontal-tension / ground-length guess.  The actual solve below is the
  # multi-segment solution.
  seed_U = 0.25 * line_length_main
  seed_segments, _ = make_segments(seed_U)
  if seed_segments is None:
    return {"success": False, "status": "FAIL",
            "msg": "No valid suspended segment lengths remain for the supplied H16/H19/H21 values."}

  W_eq = sum(seg["L"] * seg["w"] for seg in seed_segments) / line_length_main
  EA_eq = line_length_main / sum(seg["L"] / seg["EA"] for seg in seed_segments)

  try:
    HA_eq, VA_eq, _, VF_eq, _ = mc.catenary(
      XF=xf_net,
      ZF=zf,
      L=line_length_main,
      EA=EA_eq,
      W=W_eq,
      depth=water_depth,
    )
    H_guess = max(abs(HA_eq), 10.0)
    U_guess = (
      abs(VF_eq) / max(W_eq, 1e-12)
      if VA_eq <= 0
      else 0.25 * line_length_main
    )
  except Exception:
    H_guess, U_guess = 500000.0, 0.25 * line_length_main

  lower_bound_u = 0.0
  upper_bound_u = line_length_main - 1e-6
  U_guess = float(np.clip(U_guess, lower_bound_u + 1e-6, upper_bound_u))

  def residual(vars):
    H_val, U_val = vars
    segments_val, _ = make_segments(U_val)
    if segments_val is None:
      return [1e6, 1e6]

    # The optimiser only needs the end point.  Two samples per segment give
    # exactly the same end-point equations as a dense profile, while avoiding
    # thousands of short-lived arrays across the multi-start solve.  The final
    # 500/180-point evaluations below remain unchanged for validation/display.
    x_arr, z_arr, _, _, _, joints_local = eval_multisegment_profile(
      H_val, U_val, segments_val, num_points=2
    )
    if x_arr.size < 2:
      return [1e6, 1e6]

    return [
      x_arr[-1] - xf_net,
      z_arr[-1] - zf,
    ]

  # IMPORTANT:
  # H21 is a segment-placement constraint, not a third nonlinear geometry
  # variable.  H21 fixes the suspended material length split, while H and U
  # are solved so the resulting *physical curve* closes between TDP and
  # fairlead.  This lets the suspended-line angle and grounded ratio change
  # naturally as line length/H16 change.
  _u_seeds = np.unique(np.clip(
    np.concatenate((
      [U_guess],
      line_length_main * np.array(
        [0.01, 0.03, 0.05, 0.10, 0.15, 0.20, 0.30, 0.40, 0.50]
      ),
    )),
    lower_bound_u + 1e-6,
    upper_bound_u,
  ))

  _solutions = []
  for _u0 in _u_seeds:
    try:
      _trial = least_squares(
        residual,
        [max(float(H_guess), 10.0), float(_u0)],
        bounds=([10.0, lower_bound_u], [np.inf, upper_bound_u]),
        xtol=1e-11,
        ftol=1e-11,
        gtol=1e-11,
        max_nfev=2500,
      )
      if np.all(np.isfinite(_trial.x)) and np.isfinite(_trial.cost):
        _solutions.append(_trial)
    except Exception:
      continue

  if not _solutions:
    return {"success": False, "status": "FAIL",
            "msg": "Semi-Taut solver failed to close the H/U solution for the requested H16/H19/H21 configuration."}

  # Prefer genuine converged endpoint solutions, then the smallest grounded
  # length.  H/U determine the suspended geometry and ground ratio.
  _solutions.sort(key=lambda r: (float(r.cost), float(r.x[1])))
  res = _solutions[0]
  H_opt, U_opt = map(float, res.x)

  segments, L_suspended = make_segments(U_opt)
  if segments is None:
    return {"success": False, "status": "FAIL",
            "msg": "Semi-Taut solution has no valid suspended section."}

  # Final high-resolution geometry.
  _x_chk, _z_chk, _, _, _, _joints_chk = eval_multisegment_profile(
    H_opt, U_opt, segments, num_points=500
  )
  _endpoint_err = float(np.hypot(
    _x_chk[-1] - xf_net,
    _z_chk[-1] - zf,
  ))
  if _endpoint_err > 2e-3:
    return {"success": False, "status": "FAIL",
            "msg": (
              f"Semi-Taut line did not geometrically close after solving "
              f"ground ratio/angle (endpoint error={_endpoint_err:.4f} m)."
            )}

  # Re-evaluate using the standard-resolution profile used by the rest of the
  # application so the returned tensions and joint coordinates are consistent
  # with the rendered geometry.
  H_opt, U_opt = float(H_opt), float(U_opt)

  x_net, z_net, VA, VF, L_bot, joints = eval_multisegment_profile(
      H_opt, U_opt, segments, num_points=180
  )

# The three suspended sections now exactly follow H16:
  # U_opt is the grounded portion of the lower chain.  The physical lower
  # chain segment includes BOTH its grounded and suspended portions.
  L_bottom_chain_grounded = max(float(U_opt), 0.0)
  L_bottom_chain_total = float(segments[0]["L"])
  L_bottom_chain_suspended = max(
      L_bottom_chain_total - L_bottom_chain_grounded, 0.0
  )
  L_rope = float(segments[1]["L"])
  L_top_chain = float(segments[2]["L"])
  # Keep the legacy dashboard variable L_bot as the actual grounded length.
  L_bot = L_bottom_chain_grounded
  ground_ratio = float(L_bottom_chain_grounded / max(line_length_main, 1e-12))
  X_td_net = float(L_bot * (1.0 + H_opt / EA_chain)) if L_bot > 0 else 0.0
  suspended_line_angle_deg = calculate_suspended_line_angle_deg(
      water_depth, fairlead_draft, xf_net, X_td_net
  )
  ground_line_low_fail = ground_ratio < GROUND_LINE_MIN_RATIO
  ground_line_high_fail = (
      ground_ratio > GROUND_LINE_MAX_RATIO
      and suspended_line_angle_deg >= GROUND_LINE_VERTICAL_ANGLE_DEG
  )
  ground_line_fail = ground_line_low_fail or ground_line_high_fail
  ground_line_warning = False
  ground_line_warning_text = ""
  if ground_line_low_fail:
    ground_line_fail_msg = (
        f"Insufficient ground line: {ground_ratio*100.0:.1f}% of the net/main line is grounded "
        f"(minimum={GROUND_LINE_MIN_RATIO*100.0:.0f}%)."
    )
  elif ground_line_high_fail:
    ground_line_fail_msg = (
        f"Excessive ground line with near-vertical suspended section: "
        f"{ground_ratio*100.0:.1f}% grounded and suspended-line angle="
        f"{suspended_line_angle_deg:.1f}° (vertical threshold={GROUND_LINE_VERTICAL_ANGLE_DEG:.1f}°, "
        f"ground threshold={GROUND_LINE_MAX_RATIO*100.0:.0f}%)."
    )
  else:
    ground_line_fail_msg = ""
  # Pre-compute diagnostic tension/MBL/FoS quantities before any ground-line
  # or taut-section failure return.  This keeps the Failure Criteria report
  # populated even when the configuration fails before the normal PASS return.
  T_fairlead = float(np.hypot(H_opt, VF))
  T_padeye = float(np.hypot(H_opt, VA))
  T_max = max(T_fairlead, T_padeye)
  segment_checks = calculate_segment_tension_checks(H_opt, U_opt, segments)
  chain_checks = [c for c in (segment_checks[0], segment_checks[2]) if np.isfinite(c.get("T_max", np.nan))]
  rope_check = segment_checks[1]
  chain_tension_max = max([c["T_max"] for c in chain_checks], default=0.0)
  rope_tension_max = float(rope_check["T_max"]) if np.isfinite(rope_check.get("T_max", np.nan)) else float("inf")
  chain_grade, chain_mbl, _ = select_optimal_chain_grade(chain_d, chain_tension_max, min_fos=MIN_LINE_MBL_FOS)
  rope_mbl = calculate_rope_mbl(rope_d, "Polyester")
  chain_fos = chain_mbl / max(chain_tension_max, 1e-9)
  rope_fos = rope_mbl / max(rope_tension_max, 1e-9)
  anchor_results = evaluate_anchor_capacities(anchor_db, [primary_anc], H_opt, VA)

  if ground_line_fail:
    return {
        "success": False,
        "status": "FAIL",
        "msg": ground_line_fail_msg,
        "ground_ratio": ground_ratio,
        "suspended_line_angle_deg": suspended_line_angle_deg,
        "min_geometric_length": min_geometric_length,
        "excess_line": excess_line,
        "L_bot": L_bottom_chain_grounded, "L_sus": L_suspended,
        "grounded_surface_length": L_bottom_chain_grounded,
        "suspended_surface_length": L_suspended,
        "HA": H_opt, "VA": VA, "HF": H_opt, "VF": VF,
        "T_padeye": T_padeye, "T_fairlead": T_fairlead, "T_max": T_max,
        "line_fos_chain": chain_fos, "line_fos_rope": rope_fos,
        "line_fos_fairlead": chain_mbl / max(T_fairlead, 1e-9),
        "line_fos_padeye": chain_mbl / max(T_padeye, 1e-9),
        "chain_grade": chain_grade, "rope_material": "Polyester",
        "chain_d": chain_d, "rope_d": rope_d, "system_type": "Semi-Taut",
        "anchor_results": anchor_results, "primary_anc": primary_anc,
        "taut_max_tension": rope_tension_max if np.isfinite(rope_tension_max) else np.nan,
        "diagnostic_note": "Tension, MBL/FoS and anchor values shown are diagnostic values from the solved configuration before the ground-line failure criterion was applied.",
    }

  # -----------------------------------------------------------------------
  # Taut/synthetic-rope section: explicit no-slack check
  # -----------------------------------------------------------------------
  # The rope must be in tension at EVERY point along its length.  Checking
  # only the two ends is insufficient because the vertical component V can
  # pass through zero inside the rope, which is where resultant tension is
  # actually minimised.
  #
  # For this solver the horizontal component H_opt is constant through the
  # rope and the vertical component varies linearly with arc length:
  #
  #     V(s) = V_start + w_rope * s
  #
  # Therefore the minimum resultant tension occurs where V(s) is closest to
  # zero, or at one of the rope ends when zero is outside the interval.
  V_rope_start = float(segments[0]["w"]) * L_bottom_chain_suspended
  V_rope_end = V_rope_start + float(segments[1]["w"]) * L_rope

  if L_rope <= 0.0 or not np.isfinite(L_rope):
    return {
        "success": False, "status": "FAIL",
        "msg": "Taut section has zero/invalid length.",
        "HA": H_opt, "VA": VA, "HF": H_opt, "VF": VF,
        "T_padeye": T_padeye, "T_fairlead": T_fairlead, "T_max": T_max,
        "line_fos_chain": chain_fos, "line_fos_rope": rope_fos,
        "chain_grade": chain_grade, "rope_material": "Polyester",
        "chain_d": chain_d, "rope_d": rope_d, "system_type": "Semi-Taut",
        "anchor_results": anchor_results, "ground_ratio": ground_ratio,
        "grounded_surface_length": L_bottom_chain_grounded, "suspended_surface_length": L_suspended,
        "min_geometric_length": min_geometric_length, "excess_line": excess_line,
    }

  if not np.isfinite(H_opt) or H_opt <= 0.0:
    return {
        "success": False, "status": "FAIL",
        "msg": (f"Taut section is slack: horizontal tension H={H_opt:.2f} N is not positive."),
        "HA": H_opt, "VA": VA, "HF": H_opt, "VF": VF,
        "T_padeye": T_padeye, "T_fairlead": T_fairlead, "T_max": T_max,
        "line_fos_chain": chain_fos, "line_fos_rope": rope_fos,
        "chain_grade": chain_grade, "rope_material": "Polyester",
        "chain_d": chain_d, "rope_d": rope_d, "system_type": "Semi-Taut",
        "anchor_results": anchor_results, "taut_horizontal_tension": H_opt,
        "taut_max_tension": rope_tension_max, "ground_ratio": ground_ratio,
        "grounded_surface_length": L_bottom_chain_grounded, "suspended_surface_length": L_suspended,
        "min_geometric_length": min_geometric_length, "excess_line": excess_line,
    }

  w_rope_section = float(segments[1]["w"])

  # Candidate location of the minimum resultant tension where V=0.
  # Because V varies monotonically through the rope, this is the global
  # minimum of sqrt(H^2 + V^2).
  if w_rope_section > 0.0:
    s_at_min = np.clip(
        -V_rope_start / w_rope_section,
        0.0,
        L_rope,
    )
  else:
    # Fallback for an unusual zero/negative submerged-weight rope.
    s_at_min = 0.0 if abs(V_rope_start) <= abs(V_rope_end) else L_rope

  V_at_min = V_rope_start + w_rope_section * s_at_min
  taut_min_tension = float(np.hypot(H_opt, V_at_min))

  if (
      not np.isfinite(taut_min_tension)
      or taut_min_tension < SEMI_TAUT_MIN_TAUT_TENSION
  ):
    return {
        "success": False,
        "status": "FAIL",
        "msg": (
            f"Taut section is slack/under-tensioned: minimum tension="
            f"{taut_min_tension:.2f} N at s={s_at_min:.2f} m "
            f"(H={H_opt:.2f} N)."
        ),
        "HA": H_opt, "VA": VA, "HF": H_opt, "VF": VF,
        "T_padeye": T_padeye, "T_fairlead": T_fairlead, "T_max": T_max,
        "line_fos_chain": chain_fos, "line_fos_rope": rope_fos,
        "chain_grade": chain_grade, "rope_material": "Polyester",
        "chain_d": chain_d, "rope_d": rope_d, "system_type": "Semi-Taut",
        "anchor_results": anchor_results, "taut_min_tension": taut_min_tension,
        "taut_max_tension": np.nan, "taut_horizontal_tension": H_opt,
        "ground_ratio": ground_ratio, "grounded_surface_length": L_bottom_chain_grounded,
        "suspended_surface_length": L_suspended, "min_geometric_length": min_geometric_length,
        "excess_line": excess_line,
    }

  # Full tension profile check across the complete synthetic-rope section.
  # This is the governing no-slack check: every point must remain in tension
  # and above the configured minimum.
  s_check = np.linspace(0.0, L_rope, TENSION_PROFILE_POINTS)
  V_check = V_rope_start + w_rope_section * s_check
  T_check = np.hypot(H_opt, V_check)
  i_min = int(np.argmin(T_check))
  i_max = int(np.argmax(T_check))
  if (
      not np.all(np.isfinite(T_check))
      or np.any(T_check < SEMI_TAUT_MIN_TAUT_TENSION)
  ):
    return {
        "success": False,
        "status": "FAIL",
        "msg": (
            "Taut section contains a portion below the minimum tension "
            f"threshold of {SEMI_TAUT_MIN_TAUT_TENSION:.2f} N "
            f"(minimum={T_check[i_min]:.2f} N at s={s_check[i_min]:.2f} m)."
        ),
        "HA": H_opt, "VA": VA, "HF": H_opt, "VF": VF,
        "T_padeye": T_padeye, "T_fairlead": T_fairlead, "T_max": T_max,
        "line_fos_chain": chain_fos, "line_fos_rope": rope_fos,
        "chain_grade": chain_grade, "rope_material": "Polyester",
        "chain_d": chain_d, "rope_d": rope_d, "system_type": "Semi-Taut",
        "anchor_results": anchor_results, "taut_min_tension": float(T_check[i_min]),
        "taut_max_tension": float(T_check[i_max]), "taut_horizontal_tension": H_opt,
        "ground_ratio": ground_ratio, "grounded_surface_length": L_bottom_chain_grounded,
        "suspended_surface_length": L_suspended, "min_geometric_length": min_geometric_length,
        "excess_line": excess_line,
    }
  taut_max_tension = float(T_check[i_max])
  taut_min_tension_profile = float(T_check[i_min])

  # Segment MBL checks were pre-computed above so failed configurations retain
  # the same diagnostic values in the Failure Criteria report.
  if chain_fos < MIN_LINE_MBL_FOS or rope_fos < MIN_LINE_MBL_FOS:
    return {
        "success": False, "status": "FAIL",
        "msg": (f"Segment MBL screening failed: chain FoS={chain_fos:.2f}, rope FoS={rope_fos:.2f} (< {MIN_LINE_MBL_FOS:.2f})."),
        "HA": H_opt, "VA": VA, "HF": H_opt, "VF": VF,
        "T_padeye": T_padeye, "T_fairlead": T_fairlead, "T_max": T_max,
        "line_fos_chain": chain_fos, "line_fos_rope": rope_fos,
        "chain_grade": chain_grade, "rope_material": "Polyester",
        "chain_d": chain_d, "rope_d": rope_d, "system_type": "Semi-Taut",
        "anchor_results": anchor_results, "taut_min_tension": taut_min_tension_profile,
        "taut_max_tension": taut_max_tension, "taut_horizontal_tension": H_opt,
        "ground_ratio": ground_ratio, "grounded_surface_length": L_bottom_chain_grounded,
        "suspended_surface_length": L_suspended, "min_geometric_length": min_geometric_length,
        "excess_line": excess_line,
    }

  # Convert the solved profile into dashboard coordinates.
  x_plot = x_net + x_sub_span
  z_plot = z_net - water_depth
  geometry_ok, geometry_msg = validate_mooring_geometry(
      x_net, z_net, xf_net, zf, expected_start=(0.0, 0.0), label="Semi-Taut"
  )
  if not geometry_ok:
    return {
        "success": False, "status": "FAIL", "msg": geometry_msg,
        "HA": H_opt, "VA": VA, "HF": H_opt, "VF": VF,
        "T_padeye": T_padeye, "T_fairlead": T_fairlead, "T_max": T_max,
        "line_fos_chain": chain_fos, "line_fos_rope": rope_fos,
        "chain_grade": chain_grade, "rope_material": "Polyester",
        "chain_d": chain_d, "rope_d": rope_d, "system_type": "Semi-Taut",
        "anchor_results": anchor_results, "taut_min_tension": taut_min_tension_profile,
        "taut_max_tension": taut_max_tension, "taut_horizontal_tension": H_opt,
        "ground_ratio": ground_ratio, "grounded_surface_length": L_bottom_chain_grounded,
        "suspended_surface_length": L_suspended, "min_geometric_length": min_geometric_length,
        "excess_line": excess_line,
    }
  x1_n, z1_n = joints[1][0] + x_sub_span, joints[1][1] - water_depth
  x2_n, z2_n = joints[2][0] + x_sub_span, joints[2][1] - water_depth
  X_td = X_td_net + x_sub_span if L_bot > 0 else None

  # For H19=FALSE the physical Bottom Chain / Rope Joint must lie above the
  # TDP and the rendered bottom-chain segment must therefore have positive
  # suspended length.  Reject a numerically plausible solution that collapses
  # this connection.
  if not use_equal_chain_split:
    if X_td is None or x1_n <= float(X_td) + 1e-5:
      return {
          "success": False, "status": "FAIL",
          "msg": (
              "H21 Lower Joint could not be connected to the suspended bottom-chain: "
              f"TDP={float(X_td) if X_td is not None else float('nan'):.3f} m, "
              f"Lower Joint={x1_n:.3f} m."
          ),
      }

  T_fairlead = float(np.hypot(H_opt, VF))
  T_padeye = float(np.hypot(H_opt, VA))
  anchor_results = evaluate_anchor_capacities(
      anchor_db, [primary_anc], H_opt, VA
  )
  if anchor_results.get(primary_anc, {}).get("status") == "FAIL":
    return {
        "success": False, "status": "FAIL", "msg": "Geotechnical capacity failed.",
        "HA": H_opt, "VA": VA, "HF": H_opt, "VF": VF,
        "T_padeye": T_padeye, "T_fairlead": T_fairlead, "T_max": T_max,
        "line_fos_chain": chain_fos, "line_fos_rope": rope_fos,
        "chain_grade": chain_grade, "rope_material": "Polyester",
        "chain_d": chain_d, "rope_d": rope_d, "system_type": "Semi-Taut",
        "anchor_results": anchor_results, "taut_min_tension": taut_min_tension_profile,
        "taut_max_tension": taut_max_tension, "taut_horizontal_tension": H_opt,
        "ground_ratio": ground_ratio, "grounded_surface_length": L_bottom_chain_grounded,
        "suspended_surface_length": L_suspended, "min_geometric_length": min_geometric_length,
        "excess_line": excess_line,
    }

  # Physical-length closure check: subsurface length + grounded bottom chain
  # + suspended bottom chain + synthetic rope + top chain must equal the
  # specified total line length (within numerical tolerance).
  _constructed_main = (
      float(L_bottom_chain_total) + float(L_rope) + float(L_top_chain)
  )
  _length_error = float(_constructed_main - line_length_main)
  if abs(_length_error) > max(1e-6, 1e-6 * max(abs(line_length_main), 1.0)):
    return {
        "success": False,
        "status": "FAIL",
        "msg": (
            f"Semi-Taut line-length closure failed: main={_constructed_main:.6f} m, "
            f"required={line_length_main:.6f} m, error={_length_error:.6e} m."
        ),
    }

  return {
      "success": True,
      "status": "PASS",
      "x_plot": x_plot,
      "z_plot": z_plot,
      "xf": xf,
      "water_depth": water_depth,
      "fairlead_draft": fairlead_draft,
      "anchor_radius": anchor_radius,
      "line_length": line_length,
      "num_lines": num_lines,
      "num_turbines": num_turbines,
      "farm_area": farm_area,
      "buffer_zone": buffer_zone,
      "L_sus": line_length_main - L_bot,
      "L_bot": L_bot,
      "L_bottom_chain_grounded": L_bottom_chain_grounded,
      "L_bottom_chain_suspended": L_bottom_chain_suspended,
      "suspended_bottom_chain_percentage": (
          100.0 * L_bottom_chain_suspended / max(L_suspended, 1e-12)
      ),
      "top_chain_percentage": (
          100.0 * L_top_chain / max(L_suspended, 1e-12)
      ),
      "taut_percentage": 100.0 * target_taut,
      "upper_equals_lower": use_equal_chain_split,
      "lower_joint_pos": joint_fraction,
      "lower_joint_position_fraction": joint_fraction if not use_equal_chain_split else None,
      "solved_bottom_chain_suspended_fraction": (((1.0 - target_taut) / 2.0) if use_equal_chain_split else (1.0 - joint_fraction)),
      "taut_min_tension": taut_min_tension_profile,
      "taut_min_tension_location": float(s_check[i_min]),
      "taut_max_tension": taut_max_tension,
      "taut_horizontal_tension": H_opt,
      "taut_min_tension_pass": True,
      "taut_tension_profile_points": int(TENSION_PROFILE_POINTS),
      "HF": H_opt,
      "VF": VF,
      "HA": H_opt,
      "VA": VA,
      "T_fairlead": T_fairlead,
      "T_padeye": T_padeye,
      "T_max": max(T_fairlead, T_padeye),
      "line_fos_chain": chain_fos,
      "line_fos_rope": rope_fos,
      "required_line_fos": MIN_LINE_MBL_FOS,
      "segment_tension_checks": segment_checks,
      "padeye_angle_deg": padeye_angle_deg,
      "padeye_z": z_anchor,
      "padeye_position_fraction": (padeye_params or {}).get(primary_anc, {}).get("position_fraction", None),
      "X_td": X_td,
      "title": f"Semi-Taut Analysis (Multi-Segment) - {primary_anc}",
      "center_lat": center_lat,
      "center_lon": center_lon,
      "sub_x": sub_x,
      "sub_z": sub_z,
      "L_sub": L_sub,
      "min_geometric_length": min_geometric_length,
      "min_geometric_main_length": min_geometric_main,
      "excess_line": excess_line,
      "ground_ratio": ground_ratio,
      "ground_line_warning": False,
      "ground_line_warning_text": "",
      "ground_line_min_ratio": GROUND_LINE_MIN_RATIO,
      "ground_line_max_ratio": GROUND_LINE_MAX_RATIO,
      "suspended_line_angle_deg": suspended_line_angle_deg,
      "ground_line_vertical_angle_deg": GROUND_LINE_VERTICAL_ANGLE_DEG,
      "grounded_surface_length": L_bottom_chain_grounded,
      "suspended_surface_length": L_suspended,
      "rope_nodes": (x1_n, z1_n, x2_n, z2_n),
      "joints": [(x1_n, z1_n), (x2_n, z2_n)],
      "L_taut": L_rope,
      "L_top_chain": L_top_chain,
      "L_bot_chain": L_bottom_chain_total,
      "s_rope_start": L_bottom_chain_suspended,
      "s_rope_end": L_bottom_chain_suspended + L_rope,
      "line_length_main": line_length_main,
      "rope_d": rope_d,
      "chain_d": chain_d,
      "system_type": "Semi-Taut",
      "chain_grade": chain_grade,
      "rope_material": "Polyester",
      "anchor_results": anchor_results,
      "anchor_width": a_width,
      "anchor_height": a_height,
      "anchor_depth": a_depth,
      "primary_anc": primary_anc,
  }



def get_taut_effective_penetration_depth(anchor_depth, anchor_height):
  """Return the physical maximum penetration for a taut top/head anchor.

  Taut anchors are modelled with the padeye/head at the seabed surface and
  the physical anchor length extending downward. Therefore the actual bottom
  of the anchor is the maximum penetration depth, regardless of a separate
  nominal penetration-depth input.
  """
  depth_in = float(anchor_depth or 0.0)
  length_in = float(anchor_height or 0.0)
  if not np.isfinite(length_in) or length_in <= 0.0:
    return max(depth_in, 0.0)
  return length_in


def evaluate_taut_anchor(
    water_depth,
    anchor_radius,
    line_d,
    fairlead_radius,
    fairlead_draft,
    num_turbines,
    farm_area,
    num_lines,
    center_lat,
    center_lon,
    buffer_zone,
    anchor_db,
    primary_anc,
    seabed_z=None,
):
  a_width, a_height, a_depth_input = get_anchor_dimensions(anchor_db, primary_anc)
  # For taut systems the padeye is at the anchor head/top, so the physical
  # anchor length itself defines the bottom and therefore the true maximum
  # penetration depth. The nominal Excel penetration-depth field is retained
  # as an input for reporting/backward compatibility but is not used to move
  # the taut padeye below the seabed.
  a_depth = get_taut_effective_penetration_depth(a_depth_input, a_height)

  # 1. Padeye/head is exactly at the seabed surface.
  r_padeye = anchor_radius - (a_width / 2.0)
  seabed_z = (
      float(seabed_z)
      if seabed_z is not None and np.isfinite(float(seabed_z))
      else -float(water_depth)
  )
  z_anchor = seabed_z

  xf, zf = abs(r_padeye - fairlead_radius), fairlead_draft - z_anchor

  # 2. Calculate line length dynamically in Python
  # Find the direct straight-line (Euclidean) distance from anchor padeye to fairlead
  direct_span = np.sqrt(xf**2 + zf**2)

  # Set unstressed line length slightly shorter than the direct span (e.g., 99.5%)
  # so that it stretches taut under operational tension.
  line_length = direct_span * 0.995

  W, EA = abs(
      np.pi * (line_d / 2.0) ** 2 * (970.0 - 1025.0) * 9.81
  ), np.pi * (line_d / 2.0) ** 2 * 50e9

  try:
    HA, VA, HF, VF, info = mc.catenary(
        XF=xf, ZF=zf, L=line_length, EA=EA, W=W, depth=water_depth
    )

    T_max, rope_mbl = np.sqrt(HF**2 + VF**2), calculate_rope_mbl(
        line_d, "HMPE"
    )

    # Full taut-line tension profile. For a uniform distributed load, the
    # horizontal component H is constant and the vertical component varies
    # linearly along the line. Sample the complete line rather than checking
    # only fairlead/anchor tensions.
    s_taut = np.linspace(0.0, line_length, TENSION_PROFILE_POINTS)
    V_taut = VA + W * s_taut
    T_taut = np.hypot(HA, V_taut)
    if not np.all(np.isfinite(T_taut)):
      return {
          "success": False,
          "status": "FAIL",
          "msg": "Taut line tension profile contains non-finite values.",
          "line_length": line_length,
          "min_geometric_length": direct_span,
          "excess_line": line_length - direct_span,
          "xf": xf, "z_anchor_custom": z_anchor,
          "fairlead_draft": fairlead_draft,
          "rope_d": line_d, "rope_material": "HMPE",
          "HA": HA, "VA": VA, "HF": HF, "VF": VF,
      }
    i_tmin = int(np.argmin(T_taut))
    i_tmax = int(np.argmax(T_taut))
    taut_line_min_tension = float(T_taut[i_tmin])
    taut_line_max_tension = float(T_taut[i_tmax])
    if taut_line_min_tension < TAUT_MIN_TENSION:
      return {
          "success": False,
          "status": "FAIL",
          "msg": (
              f"Taut line is under-tensioned: minimum tension="
              f"{taut_line_min_tension:.2f} N at s={s_taut[i_tmin]:.2f} m "
              f"(< {TAUT_MIN_TENSION:.2f} N)."
          ),
          "line_length": line_length, "min_geometric_length": direct_span,
          "excess_line": line_length - direct_span, "xf": xf,
          "z_anchor_custom": z_anchor, "fairlead_draft": fairlead_draft,
          "rope_d": line_d, "rope_material": "HMPE",
          "HA": HA, "VA": VA, "HF": HF, "VF": VF,
          "T_max": taut_line_max_tension, "taut_min_tension": taut_line_min_tension,
          "taut_max_tension": taut_line_max_tension,
      }
    required_rope_mbl = taut_line_max_tension * MIN_LINE_MBL_FOS
    if rope_mbl < required_rope_mbl:
      return {
          "success": False,
          "status": "FAIL",
          "msg": (
              f"Taut line MBL/FoS screening failed: max tension="
              f"{taut_line_max_tension:.2f} N, MBL={rope_mbl:.2f} N, "
              f"FoS={rope_mbl / max(taut_line_max_tension, 1e-9):.2f} "
              f"(< {MIN_LINE_MBL_FOS:.2f})."
          ),
          "line_length": line_length, "min_geometric_length": direct_span,
          "excess_line": line_length - direct_span, "xf": xf,
          "z_anchor_custom": z_anchor, "fairlead_draft": fairlead_draft,
          "rope_d": line_d, "rope_material": "HMPE",
          "HA": HA, "VA": VA, "HF": HF, "VF": VF,
          "T_max": taut_line_max_tension, "taut_min_tension": taut_line_min_tension,
          "taut_max_tension": taut_line_max_tension, "line_fos": rope_mbl / max(taut_line_max_tension, 1e-9),
      }

    # DEA uplift is an anchor-specific failure, but the line tension state
    # above is still a valid calculated result.  Evaluate the DEA condition
    # after the full taut profile/line-FoS calculation so the report can show:
    #   - Line maximum tension: evaluated
    #   - Rope MBL / required BL: evaluated
    #   - Line FoS: evaluated
    #   - Anchor FoS: separate / not evaluated unless the anchor capacity model
    #     actually returned a finite FoS.
    # Positive VA is uplift at the DEA padeye and must fail the anchor.
    if primary_anc == "DEA" and VA > 1e-6:
      t_padeye = float(np.hypot(HA, VA))
      t_fairlead = float(np.hypot(HF, VF))
      return {
          "success": False,
          "status": "FAIL",
          "msg": (f"DEA uplift failure: vertical anchor tension VA={VA/1000.0:.3f} kN "
                  f"(must be <= 0 kN). Horizontal anchor tension H={HA/1000.0:.3f} kN."),
          "line_length": line_length,
          "min_geometric_length": line_length,
          "direct_geometric_span": direct_span,
          "excess_line": 0.0,
          "xf": xf, "z_anchor_custom": z_anchor,
          "fairlead_draft": fairlead_draft,
          "rope_d": line_d, "rope_material": "HMPE",
          "HA": HA, "VA": VA, "HF": HF, "VF": VF,
          "T_padeye": t_padeye, "T_fairlead": t_fairlead,
          "T_max": taut_line_max_tension,
          "taut_min_tension": taut_line_min_tension,
          "taut_max_tension": taut_line_max_tension,
          "line_fos": rope_mbl / max(taut_line_max_tension, 1e-9),
          "diagnostic_note": (
              "DEA uplift failed after the full taut tension profile and line FoS "
              "were calculated. Anchor FoS is a separate criterion and was not "
              "evaluated by the geotechnical capacity routine."
          ),
      }

    if T_max >= rope_mbl:
      return {
          "success": False,
          "status": "FAIL",
          "msg": "Tension exceeds rope MBL.",
          "line_length": line_length, "min_geometric_length": direct_span,
          "excess_line": line_length - direct_span, "xf": xf,
          "z_anchor_custom": z_anchor, "fairlead_draft": fairlead_draft,
          "rope_d": line_d, "rope_material": "HMPE",
          "HA": HA, "VA": VA, "HF": HF, "VF": VF,
          "T_max": T_max, "taut_min_tension": taut_line_min_tension,
          "taut_max_tension": taut_line_max_tension, "line_fos": rope_mbl / max(taut_line_max_tension, 1e-9),
      }

    anchor_results = evaluate_anchor_capacities(
        anchor_db, [primary_anc], HA, VA
    )
    if anchor_results.get(primary_anc, {}).get("status") == "FAIL":
      return {
          "success": False,
          "status": "FAIL",
          "msg": "Geotechnical capacity failed.",
          "line_length": line_length, "min_geometric_length": line_length,
          "direct_geometric_span": direct_span, "excess_line": 0.0,
          "xf": xf, "z_anchor_custom": z_anchor, "fairlead_draft": fairlead_draft,
          "rope_d": line_d, "rope_material": "HMPE",
          "HA": HA, "VA": VA, "HF": HF, "VF": VF,
          "T_max": taut_line_max_tension, "taut_min_tension": taut_line_min_tension,
          "taut_max_tension": taut_line_max_tension,
          "line_fos": rope_mbl / max(taut_line_max_tension, 1e-9),
          "anchor_results": anchor_results,
      }

    a, X_v = max(abs(HF), 1.0) / max(W, 1e-3), -(
        max(abs(HF), 1.0) / max(W, 1e-3)
    ) * np.arcsinh(abs(VA) / max(abs(HF), 1.0))
    x_plot = np.linspace(0, xf, 500)
    z_plot = np.array([
        z_anchor + a * (np.cosh((x - X_v) / a) - np.cosh(-X_v / a))
        for x in x_plot
    ])
    z_plot[-1] = fairlead_draft

    return {
        "success": True,
        "status": "PASS",
        "x_plot": x_plot,
        "z_plot": z_plot,
        "xf": xf,
        "water_depth": water_depth,
        "fairlead_draft": fairlead_draft,
        "anchor_radius": anchor_radius,
        "line_length": line_length,
        "min_geometric_length": line_length,  # Taut design minimum is the generated optimal taut length.
        "min_geometric_main_length": line_length,
        "direct_geometric_span": direct_span,
        "excess_line": 0.0,
        "num_lines": num_lines,
        "num_turbines": num_turbines,
        "farm_area": farm_area,
        "buffer_zone": buffer_zone,
        "L_sus": line_length,
        "L_bot": 0.0,
        "HF": HF,
        "VF": VF,
        "HA": HA,
        "VA": VA,
        "taut_min_tension": taut_line_min_tension,
        "taut_min_tension_location": float(s_taut[i_tmin]),
        "taut_max_tension": taut_line_max_tension,
        "taut_tension_profile_points": int(TENSION_PROFILE_POINTS),
        "taut_min_tension_pass": True,
        "line_fos": rope_mbl / max(taut_line_max_tension, 1e-9),
        "X_td": None,
        "title": f"Taut Analysis - {primary_anc}",
        "center_lat": center_lat,
        "center_lon": center_lon,
        "rope_d": line_d,
        "system_type": "Taut",
        "rope_material": "HMPE",
        "anchor_results": anchor_results,
        "z_anchor_custom": z_anchor,
        "seabed_z": seabed_z,
        "padeye_z": z_anchor,
        "anchor_top_z": z_anchor,
        "anchor_bottom_z": z_anchor - float(a_height),
        "anchor_width": a_width,
        "anchor_height": a_height,
        "anchor_depth": a_depth,
      "nominal_anchor_depth": a_depth_input,
        "primary_anc": primary_anc,
    }
  except Exception as e:
    return {"success": False, "status": "FAIL", "msg": str(e)}


# =========================================================================
# UPDATED EXPORT MANAGER FOR QGIS COMPATIBILITY
# =========================================================================
def export_system_to_csv(export_data, parent_figure=None):
  """Let the user choose where to save QGIS-compatible data."""

  default_filename = (
      f"qgis_export_{export_data['system_type']}_{export_data['primary_anc']}"
      .replace(" ", "_")
      .lower()
      + ".csv"
  )

  filename = _ask_to_save_file(
      parent_figure,
      title="Save QGIS Export CSV As",
      defaultextension=".csv",
      filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
      initialfile=default_filename,
  )

  if not filename:
    print("\n[INFO] Export cancelled by user.")
    return None

  try:
    with open(filename, mode="w", newline="", encoding="utf-8") as csv_file:
      writer = csv.writer(csv_file)
      # UPDATED CSV HEADER TO INCLUDE COLOR HEX
      writer.writerow([
          "Feature_Type",
          "Sub_Type",
          "Turbine_ID",
          "Line_Heading_Deg",
          "X_Coord",
          "Y_Coord",
          "Z_Coord",
          "Max_Penetration_Depth_m",
          "Anchor_Width_m",
          "Anchor_Height_m",
          "Anchor_Length_m",
          "Cylindrical_Diameter_m",
          "Cylindrical_Length_m",
          "Max_Water_Depth_m",
          "Seabed_Level_m",
          "Sea_Level_m",
          "Local_Seabed_Z_m",
          "Color_Hex",
          "Description",
      ])
      for row in export_data["rows"]:
        writer.writerow(row)
    print(
        f"\n[SUCCESS] QGIS-compatible data successfully exported to '{filename}'!"
    )
    export_data["last_exported_csv"] = filename
    return filename
  except Exception as e:
    print(f"\n[ERROR] Failed to export CSV file: {e}")
    return None


def farm_perimeter_coordinates(data, origin_x, origin_y, turbine_coords=None):
  """Return the exported farm-perimeter XY path without building CSV rows."""
  if turbine_coords is None:
    anchor_radius = float(data.get("anchor_radius", 0.0))
    num_turbines = max(0, int(data.get("num_turbines", 1)))
    buffer_zone = float(data.get("buffer_zone", 50.0))
    farm_area = data.get("farm_area")
    if num_turbines <= 0:
      return np.array([], dtype=float), np.array([], dtype=float)
    min_spacing = (2.0 * anchor_radius) + buffer_zone
    base_spacing = (
        np.sqrt((farm_area * 1e6) / num_turbines)
        if farm_area and farm_area > 0 else min_spacing
    )
    spacing = max(min_spacing, base_spacing)
    rows, cols = get_optimal_grid_dimensions(num_turbines)
    total_width_x = cols * spacing
    total_height_y = rows * spacing * 0.866
    turbine_coords = [
        (
            origin_x - (total_width_x / 2.0)
            + (col_idx + 0.5 * (row_idx % 2)) * spacing,
            origin_y - (total_height_y / 2.0) + row_idx * spacing * 0.866,
        )
        for row_idx, col_idx in (
            divmod(index, cols) for index in range(num_turbines)
        )
    ]

  if not turbine_coords:
    return np.array([], dtype=float), np.array([], dtype=float)

  anchor_radius = float(data.get("anchor_radius", 0.0))
  min_cx, max_cx = min(tc[0] for tc in turbine_coords), max(
      tc[0] for tc in turbine_coords
  )
  min_cy, max_cy = min(tc[1] for tc in turbine_coords), max(
      tc[1] for tc in turbine_coords
  )
  # This 50 m allowance is part of the exported Farm_Perimeter geometry.
  perimeter_buffer = anchor_radius + 50.0
  xmin, xmax = min_cx - perimeter_buffer, max_cx + perimeter_buffer
  ymin, ymax = min_cy - perimeter_buffer, max_cy + perimeter_buffer
  num_points_per_side = 30
  return (
      np.concatenate((
          np.linspace(xmin, xmax, num_points_per_side),
          np.full(num_points_per_side, xmax),
          np.linspace(xmax, xmin, num_points_per_side),
          np.full(num_points_per_side, xmin),
      )),
      np.concatenate((
          np.full(num_points_per_side, ymin),
          np.linspace(ymin, ymax, num_points_per_side),
          np.full(num_points_per_side, ymax),
          np.linspace(ymax, ymin, num_points_per_side),
      )),
  )


def generate_export_rows(data, origin_x, origin_y):
  x_plot = data.get("x_plot", [])
  z_plot = data.get("z_plot", [])
  xf = data.get("xf", 0.0)
  water_depth = data.get("water_depth", 0.0)
  fairlead_draft = data.get("fairlead_draft", 0.0)
  anchor_radius = data.get("anchor_radius", 0.0)
  num_lines = data.get("num_lines", 3)
  num_turbines = data.get("num_turbines", 1)
  farm_area = data.get("farm_area")
  buffer_zone = data.get("buffer_zone", 50.0)
  sub_x = data.get("sub_x")
  sub_z = data.get("sub_z")
  system_type = data.get("system_type", "Catenary")

  primary_anc = str(data.get("primary_anc", "DEA")).strip()
  anchor_depth = float(data.get("anchor_depth") or 8.0)
  anchor_height = float(data.get("anchor_height") or 3.0)
  anchor_width = float(data.get("anchor_width") or 4.0)
  z_anchor_custom = data.get("z_anchor_custom")
  X_td = data.get("X_td")
  # Bottom of anchor (maximum penetration)
  effective_penetration_depth = (
      get_taut_effective_penetration_depth(anchor_depth, anchor_height)
      if system_type == "Taut" else anchor_depth
  )
  max_penetration_z = -water_depth - effective_penetration_depth

  # Fixed cylindrical-pile geometry. For taut systems the padeye/head is at
  # the seabed and the physical anchor length defines the bottom/max penetration.
  pile_frac = float((data.get("padeye_params") or {}).get(primary_anc, {}).get("position_fraction", 0.50))
  if system_type == "Taut":
    pile_z_top = -float(water_depth)
    pile_z_bottom = pile_z_top - float(anchor_height)
    pile_z_padeye = pile_z_top
  else:
    pile_z_bottom, pile_z_top, pile_z_padeye = get_pile_anchor_vertical_geometry(
        water_depth, anchor_depth, anchor_height, padeye_fraction=pile_frac
    )
  anchor_top_z = pile_z_top

  # NEW: Extract rope nodes for Semi-Taut systems
  rope_nodes = data.get("rope_nodes")

  # Per-turbine raster seabed datum used only by the intermediate corridor
  # rendering. Physical Mooring_Line / Anchor_Body / Padeye Z_Coord values are
  # exported independently and are never replaced by this value.
  local_turbine_seabed_z = {}

  def create_row(
      feature_type, sub_type, turbine_id, heading, x, y, z, color_hex, desc
  ):
    if feature_type == "Anchor_Body":
      pen_depth = effective_penetration_depth
    else:
      pen_depth = max(0.0, -water_depth - z)
    width_val, height_val, length_val, cyl_dia_val, cyl_len_val = (
        "",
        "",
        "",
        "",
        "",
    )
    anc_type_upper = primary_anc.upper()

    if feature_type == "Anchor_Body":
      if any(
          k in anc_type_upper for k in ["DEA", "GRAVITY", "DRAG", "STEVSHARK"]
      ):
        width_val, height_val, length_val = (
            anchor_width,
            anchor_height,
            anchor_width,
        )
      elif any(
          k in anc_type_upper for k in ["DRIVEN", "DRILLED", "SUCTION", "PILE"]
      ):
        cyl_dia_val = anchor_width
        cyl_len_val = anchor_height

    return [
        feature_type,
        sub_type,
        turbine_id,
        heading,
        x,
        y,
        z,
        pen_depth,
        width_val,
        height_val,
        length_val,
        cyl_dia_val,
        cyl_len_val,
        water_depth,
        -water_depth,
        0.0,
        float(local_turbine_seabed_z.get(
            str(turbine_id), -float(water_depth)
        )),
        color_hex,
        desc,
    ]

  min_spacing = (2.0 * anchor_radius) + buffer_zone
  base_spacing = (
      np.sqrt((farm_area * 1e6) / num_turbines)
      if farm_area and farm_area > 0
      else min_spacing
  )
  spacing = max(min_spacing, base_spacing)

  rows, cols = get_optimal_grid_dimensions(num_turbines)
  total_width_x = cols * spacing
  total_height_y = rows * spacing * 0.866

  headings = generate_mooring_headings(
      num_lines,
      triad=bool(data.get("triad", False)),
      inter_arm_angle_deg=float(data.get("inter_arm_angle", 0.0)),
      angle_between_mooring_arms_deg=float(
          data.get("angle_between_mooring_arms", 360.0 / max(num_lines, 1))
      ),
  )
  base_circle_angles = np.linspace(0, 2 * np.pi, 360)
  circle_angles = np.sort(
      np.unique(np.concatenate((base_circle_angles, np.radians(headings))))
  )

  r_anchor = anchor_radius
  r_padeye = (
      anchor_radius - (anchor_width / 2.0)
      if system_type == "Taut"
      else anchor_radius - anchor_width
  )
  r_fairlead = r_padeye - xf
  r_tdp = r_padeye - X_td if X_td is not None else None
  r_ddp = (
      r_padeye - sub_x[-1] if (sub_x is not None and len(sub_x) > 0) else None
  )

  if system_type == "Taut":
    # Taut padeye/head sits at the seabed surface; anchor bottom defines max penetration.
    z_anchor_3d = -float(water_depth)
  else:
    if any(k in primary_anc.upper() for k in ("SUCTION", "DRIVEN", "DRILLED", "PILE")):
      frac = float((data.get("padeye_params") or {}).get(primary_anc, {}).get("position_fraction", 0.50))
      _, _, z_anchor_3d = get_pile_anchor_vertical_geometry(
          water_depth, anchor_depth, anchor_height, padeye_fraction=frac
      )
    else:
      z_anchor_3d = (
          sub_z[0]
          if (sub_z is not None and len(sub_z) > 0)
          else (-water_depth - anchor_depth + 0.5 * anchor_height)
      )

  turbine_coords = []
  for i in range(num_turbines):
    row_idx, col_idx = i // cols, i % cols
    cx = origin_x - (total_width_x / 2.0) + (col_idx + 0.5 * (row_idx % 2)) * spacing
    cy = origin_y - (total_height_y / 2.0) + row_idx * spacing * 0.866
    turbine_coords.append((cx, cy))

  # Per-turbine local raster seabed datum used only for the intermediate
  # square/triangle volume. This deliberately does not touch any physical
  # Mooring_Line, Anchor_Body or Padeye Z coordinate.
  raster_path = data.get("_raster_path")
  local_turbine_seabed_z = {}
  for _ti, (_cx, _cy) in enumerate(turbine_coords, start=1):
    _tz = np.nan
    if raster_path:
      try:
        _tz = float(sample_raster_seabed(raster_path, _cx, _cy))
      except Exception:
        _tz = np.nan
    if not np.isfinite(_tz):
      _tz = -float(water_depth)
    local_turbine_seabed_z[f"Turbine_{_ti}"] = float(_tz)

  # Per-line raster seabed elevations are authoritative for Taut export.
  # They are produced by evaluate_raster_configuration; if unavailable, sample
  # the configured raster directly at the actual padeye XY.
  raster_path = data.get("_raster_path")
  raster_anchor_elevations = data.get("_raster_anchor_elevations", {}) or {}

  def taut_export_seabed_z(turbine_idx, line_idx, px, py):
    key = f"{int(turbine_idx)+1}:{int(line_idx)+1}"
    cached = raster_anchor_elevations.get(key, np.nan)
    if np.isfinite(float(cached)):
      return float(cached)
    if raster_path:
      try:
        rz = float(sample_raster_seabed(raster_path, px, py))
        if np.isfinite(rz):
          return rz
      except Exception:
        pass
    return -float(water_depth)

  # Build each inverted-bridle export geometry once from the corresponding
  # per-turbine/per-line solver result.  The upper trunk and the two branches
  # must use this same source profile, otherwise their Lower Joint can be
  # sampled from different water depths.
  _bridle_export_context = {}

  def _get_inverted_bridle_export_context(
      turbine_index, line_index, heading_deg, turbine_xy,
  ):
    key = (int(turbine_index), int(line_index))
    if key in _bridle_export_context:
      return _bridle_export_context[key]

    source_data = dict(data)
    try:
      raster_results = data.get("_raster_turbine_results", []) or []
      line_set = (
          raster_results[int(turbine_index)].get("lines", [])
          if (
              int(turbine_index) < len(raster_results)
              and isinstance(raster_results[int(turbine_index)], dict)
          )
          else []
      )
      if int(line_index) < len(line_set):
        local_result = line_set[int(line_index)]
        if isinstance(local_result, dict):
          source_data.update(local_result)
    except Exception:
      pass

    source_data["_raster_path"] = raster_path
    try:
      geometry = calculate_inverted_bridle_geometry(
          source_data,
          heading_deg,
          turbine_xy,
          raster_path=raster_path,
      )
    except Exception:
      geometry = None

    context = (source_data, geometry)
    _bridle_export_context[key] = context
    return context

  export_rows = []

  edge_x, edge_y = farm_perimeter_coordinates(
      data, origin_x, origin_y, turbine_coords=turbine_coords
  )
  if edge_x.size:
    for ex, ey in zip(edge_x, edge_y):
      export_rows.append(create_row(
          "Farm_Perimeter",
          "Perimeter_Box",
          "Farm",
          "",
          ex,
          ey,
          -water_depth,
          "#800080",
          "Wind Farm Perimeter Boundary",
      ))

  for i, (cx, cy) in enumerate(turbine_coords):
    export_rows.append(create_row(
        "Turbine",
        "Location",
        f"Turbine_{i+1}",
        "",
        cx,
        cy,
        fairlead_draft,
        "#0000FF",
        f"Floating Wind Turbine {i+1} Location Symbol",
    ))

    for th in circle_angles:
      export_rows.append(create_row(
          "Boundary_Radius",
          "Mooring_Radius_Circle",
          f"Turbine_{i+1}",
          "",
          cx + r_anchor * np.cos(th),
          cy + r_anchor * np.sin(th),
          -water_depth,
          "#FF0000",
          "Mooring Radius Limit",
      ))

      if r_tdp is not None and r_tdp > 0:
        export_rows.append(create_row(
            "Boundary_Radius",
            "TDP_Radius_Circle",
            f"Turbine_{i+1}",
            "",
            cx + r_tdp * np.cos(th),
            cy + r_tdp * np.sin(th),
            -water_depth,
            "#FFD700",
            "Touchdown Point Radius Limit",
        ))

      if r_ddp is not None and r_ddp > 0:
        ddp_z_val = sub_z[-1] if sub_z is not None else -water_depth
        export_rows.append(create_row(
            "Boundary_Radius",
            "DDP_Radius_Circle",
            f"Turbine_{i+1}",
            "",
            cx + r_ddp * np.cos(th),
            cy + r_ddp * np.sin(th),
            ddp_z_val,
            "#FFA500",
            "Down Dip Point Radius Limit",
        ))

    for line_index, angle in enumerate(headings):
      angle_rad = np.radians(angle)
      ux, uy = np.cos(angle_rad), np.sin(angle_rad)
      padeye_3d_x, padeye_3d_y = cx + r_padeye * ux, cy + r_padeye * uy
      local_taut_seabed_z = (
          taut_export_seabed_z(i, int(np.argmin(np.abs(np.asarray(headings, dtype=float) - float(angle)))), padeye_3d_x, padeye_3d_y)
          if system_type == "Taut" else None
      )
      fairlead_3d_x, fairlead_3d_y = cx + r_fairlead * ux, cy + r_fairlead * uy
      vx, vy = -uy, ux

      _bridle_active = bool(data.get("inverted_bridle", False))
      _bridle_source_data = data
      _bridle_geom = None
      if _bridle_active:
        _bridle_source_data, _bridle_geom = (
            _get_inverted_bridle_export_context(
                i, line_index, angle, (cx, cy)
            )
        )

      # Use the local raster-resolved solver profile only when it produced a
      # valid bridle geometry.  Non-bridle exports retain their established
      # global solver path unchanged.
      _trunk_data = _bridle_source_data if _bridle_geom is not None else data
      _trunk_anchor_radius = float(
          _trunk_data.get("anchor_radius", anchor_radius)
      )
      _trunk_anchor_width = float(
          _trunk_data.get("anchor_width", anchor_width)
      )
      _trunk_r_padeye = _trunk_anchor_radius - (
          _trunk_anchor_width / 2.0
          if system_type == "Taut" else _trunk_anchor_width
      )

      if not bool(data.get("inverted_bridle", False)):
        anc_type_upper = primary_anc.upper()
        padeye_z = (
            local_taut_seabed_z
            if system_type == "Taut"
            else (pile_z_padeye if any(k in anc_type_upper for k in ("DRIVEN", "DRILLED")) else z_anchor_3d)
        )
        if any(
            k in anc_type_upper for k in ["DRIVEN", "DRILLED", "SUCTION", "PILE"]
        ):
          pile_radius = anchor_width / 2.0
          pile_center_offset = anchor_radius - pile_radius
          center_x, center_y = (
              cx + pile_center_offset * ux,
              cy + pile_center_offset * uy,
          )
          for th_circle in np.linspace(0, 2 * np.pi, 16, endpoint=False):
            export_rows.append(create_row(
                "Anchor_Body",
                primary_anc,
                f"Turbine_{i+1}",
                angle,
                center_x + pile_radius * np.cos(th_circle),
                center_y + pile_radius * np.sin(th_circle),
                (local_taut_seabed_z if system_type == "Taut" else anchor_top_z),
                "#800080",
                f"Anchor Body ({primary_anc} Cylinder)",
            ))
        else:
          half_w, anchor_len = anchor_width / 2.0, anchor_width
          outer_x, outer_y = cx + anchor_radius * ux, cy + anchor_radius * uy
          anchor_corners = [
              (outer_x + half_w * vx, outer_y + half_w * vy),
              (outer_x - half_w * vx, outer_y - half_w * vy),
              (
                  outer_x - half_w * vx - anchor_len * ux,
                  outer_y - half_w * vy - anchor_len * uy,
              ),
              (
                  outer_x + half_w * vx - anchor_len * ux,
                  outer_y + half_w * vy - anchor_len * uy,
              ),
          ]
          for c_x, c_y in anchor_corners:
            export_rows.append(create_row(
                "Anchor_Body",
                primary_anc,
                f"Turbine_{i+1}",
                angle,
                c_x,
                c_y,
                (local_taut_seabed_z if system_type == "Taut" else anchor_top_z),
                "#800080",
                f"Anchor Body ({primary_anc})",
            ))

        export_rows.append(create_row(
            "Anchor_Component",
            "Padeye",
            f"Turbine_{i+1}",
            angle,
            padeye_3d_x,
            padeye_3d_y,
            padeye_z,
            "#FF0000",
            "Anchor Padeye Location",
        ))
      export_rows.append(create_row(
          "Anchor_Component",
          "Fairlead",
          f"Turbine_{i+1}",
          angle,
          fairlead_3d_x,
          fairlead_3d_y,
          fairlead_draft,
          "#808080",
          "Mooring Line Fairlead Connection Point",
      ))

      if (not bool(data.get("inverted_bridle", False))) and X_td is not None and r_tdp is not None:
        export_rows.append(create_row(
            "Anchor_Component",
            "TDP",
            f"Turbine_{i+1}",
            angle,
            cx + r_tdp * ux,
            cy + r_tdp * uy,
            -water_depth,
            "#FFFF00",
            "Touchdown Point (TDP)",
        ))

      if (not bool(data.get("inverted_bridle", False))) and sub_x is not None and sub_z is not None and r_ddp is not None:
        export_rows.append(create_row(
            "Anchor_Component",
            "DDP",
            f"Turbine_{i+1}",
            angle,
            cx + r_ddp * ux,
            cy + r_ddp * uy,
            sub_z[-1],
            "#FFFF00",
            "Down Dip Point (DDP)",
        ))

      _trunk_x_plot_raw = _trunk_data.get("x_plot", x_plot)
      _trunk_z_plot_raw = _trunk_data.get("z_plot", z_plot)
      _trunk_sub_x_raw = _trunk_data.get("sub_x", sub_x)
      _trunk_sub_z_raw = _trunk_data.get("sub_z", sub_z)
      _trunk_x_plot = np.asarray(
          _trunk_x_plot_raw if _trunk_x_plot_raw is not None else [], dtype=float
      )
      _trunk_z_plot = np.asarray(
          _trunk_z_plot_raw if _trunk_z_plot_raw is not None else [], dtype=float
      )
      _trunk_sub_x = np.asarray(
          _trunk_sub_x_raw if _trunk_sub_x_raw is not None else [], dtype=float
      )
      _trunk_sub_z = np.asarray(
          _trunk_sub_z_raw if _trunk_sub_z_raw is not None else [], dtype=float
      )
      full_x, full_z = (
          np.concatenate((_trunk_sub_x, _trunk_x_plot[1:]))
          if (_trunk_sub_x.size > 0 and _trunk_sub_z.size == _trunk_sub_x.size)
          else _trunk_x_plot,
          np.concatenate((_trunk_sub_z, _trunk_z_plot[1:]))
          if (_trunk_sub_x.size > 0 and _trunk_sub_z.size == _trunk_sub_x.size)
          else _trunk_z_plot,
      )

      # AUTHORITATIVE TAUT EXPORT PROFILE
      #
      # For Taut, the physical CSV line comes only from the per-line raster
      # re-solution. The old/global Excel-depth profile is never used.
      taut_export_x = None
      taut_export_z = None
      if system_type == "Taut":
        _line_idx0 = int(np.argmin(
            np.abs(np.asarray(headings, dtype=float) - float(angle))
        ))
        _line_result = None
        try:
          _raster_results = data.get("_raster_turbine_results", []) or []
          _line_set = (
              _raster_results[i].get("lines", [])
              if i < len(_raster_results) and isinstance(_raster_results[i], dict)
              else []
          )
          if _line_idx0 < len(_line_set):
            _line_result = _line_set[_line_idx0]
        except Exception:
          pass

        if not isinstance(_line_result, dict):
          raise ValueError(
              f"Taut raster result missing for T{i+1}L{_line_idx0+1}. "
              "Run the raster-aware dashboard calculation before exporting."
          )

        # Engineering feasibility (PASS/FAIL) is separate from geometry.
        # Export/render a valid raster-resolved geometry even when the line
        # fails uplift/FoS/capacity checks, but NEVER fall back to the old
        # global Excel-depth geometry.
        taut_export_x = np.asarray(
            _line_result.get("x_plot", []), dtype=float
        )
        taut_export_z = np.asarray(
            _line_result.get("z_plot", []), dtype=float
        )

        if taut_export_x.size < 2 or taut_export_z.size != taut_export_x.size:
          raise ValueError(
              f"Taut raster geometry unavailable for T{i+1}L{_line_idx0+1}; "
              "the raster-resolved profile could not be constructed."
          )

        taut_export_x = taut_export_x.copy()
        taut_export_z = taut_export_z.copy()

        # Exact physical endpoints: local raster padeye -> fairlead.
        taut_export_x[0] = 0.0
        taut_export_z[0] = float(
            _line_result.get("seabed_z", local_taut_seabed_z)
        )
        taut_export_x[-1] = float(
            _line_result.get("xf", r_padeye - r_fairlead)
        )
        taut_export_z[-1] = float(fairlead_draft)

      # Export Joint Nodes and Split the Line Segments for Semi-Taut configurations.
      # A valid bridle uses the per-line result that produced its lower joint.
      _trunk_rope_nodes = _trunk_data.get("rope_nodes", rope_nodes)
      if system_type == "Semi-Taut" and _trunk_rope_nodes is not None:
        x1_n, z1_n, x2_n, z2_n = _trunk_rope_nodes

        # 1. Export Joint 1 (Bottom Chain to Rope)
        j1_x, j1_y = (
            cx + (_trunk_r_padeye - x1_n) * ux,
            cy + (_trunk_r_padeye - x1_n) * uy,
        )
        export_rows.append(create_row(
            "Mooring_Joint",
            "Chain_Rope_Bottom",
            f"Turbine_{i+1}",
            angle,
            j1_x,
            j1_y,
            z1_n,
            "#FF00FF",
            "Bottom Joint (Chain to Rope)",
        ))

        # 2. Export Joint 2 (Rope to Top Chain)
        j2_x, j2_y = (
            cx + (_trunk_r_padeye - x2_n) * ux,
            cy + (_trunk_r_padeye - x2_n) * uy,
        )
        export_rows.append(create_row(
            "Mooring_Joint",
            "Rope_Chain_Top",
            f"Turbine_{i+1}",
            angle,
            j2_x,
            j2_y,
            z2_n,
            "#FF00FF",
            "Top Joint (Rope to Chain)",
        ))

        # 3. Export line points. With an inverted bridle active, the original
        # lower-chain portion terminates at Joint 1 and is replaced by the two
        # bridle branches exported below.
        _joint_x = (
            float(_bridle_geom["joint_local"][0])
            if _bridle_geom is not None
            else (
                float(x1_n) if _bridle_active else None
            )
        )
        if _bridle_geom is not None:
          _joint_x_3d, _joint_y_3d, _joint_z_3d = map(
              float, _bridle_geom["joint"]
          )
          export_rows.append(create_row(
              "Mooring_Line", "Rope_Synthetic", f"Turbine_{i+1}", angle,
              _joint_x_3d, _joint_y_3d, _joint_z_3d, "#00BFFF",
              "Inverted Bridle Lower Joint to Synthetic Rope",
          ))
        for fx, fz in zip(full_x, full_z):
          if (
              _bridle_geom is not None
              and _joint_x is not None
              and fx <= _joint_x + 1e-5
          ):
            continue
          if (
              _bridle_geom is None
              and _bridle_active
              and _joint_x is not None
              and fx < _joint_x - 1e-5
          ):
            continue
          if fx <= x1_n + 1e-5:
            segment_type = "Chain_Bottom"
            segment_color = "#000000"
            desc = "Bottom Chain Section"
          elif fx <= x2_n + 1e-5:
            segment_type = "Rope_Synthetic"
            segment_color = "#00BFFF"
            desc = "Synthetic Rope Section"
          else:
            segment_type = "Chain_Top"
            segment_color = "#000000"
            desc = "Top Chain Section"
          export_rows.append(create_row(
              "Mooring_Line", segment_type, f"Turbine_{i+1}", angle,
              cx + (_trunk_r_padeye - fx) * ux,
              cy + (_trunk_r_padeye - fx) * uy, fz,
              segment_color, desc,
          ))
      else:
        # Standard fallback for Catenary & Taut lines
        line_color = "#00BFFF" if "Taut" in system_type else "#000000"
        _joint_x = float(_bridle_geom["joint_local"][0]) if _bridle_geom else None
        _export_x = taut_export_x if system_type == "Taut" else full_x
        _export_z = taut_export_z if system_type == "Taut" else full_z
        if _bridle_geom is not None:
          _joint_x_3d, _joint_y_3d, _joint_z_3d = map(
              float, _bridle_geom["joint"]
          )
          export_rows.append(create_row(
              "Mooring_Line", system_type, f"Turbine_{i+1}", angle,
              _joint_x_3d, _joint_y_3d, _joint_z_3d, line_color,
              f"{system_type} Profile (Inverted Bridle Lower Joint)",
          ))
        for fx, fz in zip(_export_x, _export_z):
          if (
              _joint_x is not None
              and _bridle_geom is not None
              and fx <= _joint_x + 1e-5
          ):
            continue
          export_rows.append(create_row(
              "Mooring_Line", system_type, f"Turbine_{i+1}", angle,
              cx + (_trunk_r_padeye - fx) * ux,
              cy + (_trunk_r_padeye - fx) * uy, fz,
              line_color, f"{system_type} Profile",
          ))

  # -----------------------------------------------------------------------
  # INVERTED BRIDLE EXPORT
  # -----------------------------------------------------------------------
  # Export both physical bridle anchors completely (body + Padeye), retain
  # importer-compatible Sub_Type="Padeye", and export each a/b branch with
  # its own ordered 3-D points.
  if bool(data.get("inverted_bridle", False)):
    for i, (cx, cy) in enumerate(turbine_coords):
      for line_idx, angle in enumerate(headings, start=1):
        _bg_data, bg = _get_inverted_bridle_export_context(
            i, line_idx - 1, angle, (cx, cy)
        )
        if not bg:
          continue

        # Build the export branches from the same per-line raster/solver
        # result that was used to locate this physical bridle geometry.  Using
        # the global ``data`` here mixed a line-specific Lower Joint/Padeye
        # with another profile's source curve, which can legitimately fail the
        # endpoint validation and abort the export callback.
        bg["profiles"] = bg.get("profiles") or build_inverted_bridle_branch_profiles(
            _bg_data, bg
        )
        jx, jy, jz = bg["joint"]
        export_rows.append(create_row(
            "Mooring_Joint", "Inverted_Bridle_Lower", f"Turbine_{i+1}",
            angle, jx, jy, jz, "#FF00FF", "Inverted Bridle Lower Joint"
        ))

        anc_type_upper = primary_anc.upper()
        is_pile = any(k in anc_type_upper for k in ("DRIVEN", "DRILLED", "SUCTION", "PILE"))
        is_box = any(k in anc_type_upper for k in ("DEA", "GRAVITY", "DRAG", "STEVSHARK"))
        branch_padeye_frac = float((data.get("padeye_params") or {}).get(primary_anc, {}).get("position_fraction", 0.50))
        pile_z_bottom_b, pile_z_top_b, _ = get_pile_anchor_vertical_geometry(
            water_depth, anchor_depth, anchor_height, padeye_fraction=branch_padeye_frac
        )

        branch_profiles = bg.get("profiles") or []
        for b_idx, a in enumerate(bg["anchors"][:2], start=1):
          bh = np.radians(float(a["heading_deg"]))
          bux, buy = np.cos(bh), np.sin(bh)
          bvx, bvy = -buy, bux
          pad_x, pad_y, pad_z = float(a["x"]), float(a["y"]), float(a["z"])

          # IMPORTANT: use the standard Padeye subtype so the QGIS importer
          # identifies these as Anchor_Component / Padeye records.
          export_rows.append(create_row(
              "Anchor_Component", "Padeye", f"Turbine_{i+1}",
              float(a["heading_deg"]), pad_x, pad_y, pad_z, "#FF0000",
              f"Inverted Bridle Anchor Padeye {b_idx} (T{i+1}L{line_idx}{'a' if b_idx == 1 else 'b'})"
          ))

          # Export the physical anchor body using the same dimensions/orientation
          # convention as the original non-bridle anchor export.
          if system_type == "Taut":
            body_z_top = pad_z
            body_z_bottom = body_z_top - anchor_height
          else:
            body_z_top = float(pile_z_top_b)
            body_z_bottom = float(pile_z_bottom_b)

          if is_pile:
            # Taut: padeye is at the head/centre of the physical anchor.
            # Catenary/Semi-Taut: retain the original side-mounted padeye
            # convention by placing the body one radius outward along heading.
            radius = anchor_width / 2.0
            if system_type == "Taut":
              body_cx = pad_x
              body_cy = pad_y
            else:
              body_cx = pad_x + radius * bux
              body_cy = pad_y + radius * buy
            for th_circle in np.linspace(0.0, 2.0*np.pi, 24, endpoint=False):
              export_rows.append(create_row(
                  "Anchor_Body", primary_anc, f"Turbine_{i+1}",
                  float(a["heading_deg"]),
                  body_cx + radius*np.cos(th_circle),
                  body_cy + radius*np.sin(th_circle),
                  body_z_top, "#800080",
                  f"Inverted Bridle Anchor Body {b_idx} ({primary_anc})"
              ))
          elif is_box:
            # Taut: centre the body on the padeye.
            # Catenary/Semi-Taut: the padeye is mounted on the side of the
            # square/DEA anchor.  The physical body therefore extends one
            # half-width outward from the padeye.  This matches the established
            # non-bridle square-anchor convention below, where the anchor centre
            # lies half an anchor width beyond the padeye.
            half_w = anchor_width / 2.0
            anchor_len = anchor_width
            if system_type == "Taut":
              outer_x = pad_x
              outer_y = pad_y
            else:
              outer_x = pad_x + half_w * bux
              outer_y = pad_y + half_w * buy
            corners = [
                (outer_x + half_w*bvx + half_w*bux, outer_y + half_w*bvy + half_w*buy),
                (outer_x - half_w*bux + half_w*bvx, outer_y - half_w*buy + half_w*bvy),
                (outer_x - half_w*bux - half_w*bvx, outer_y - half_w*buy - half_w*bvy),
                (outer_x + half_w*bux - half_w*bvx, outer_y + half_w*buy - half_w*bvy),
            ]
            for ex, ey in corners:
              export_rows.append(create_row(
                  "Anchor_Body", primary_anc, f"Turbine_{i+1}",
                  float(a["heading_deg"]), ex, ey, body_z_top, "#800080",
                  f"Inverted Bridle Anchor Body {b_idx} ({primary_anc})"
              ))
          else:
            # Generic fallback follows the same system-dependent convention.
            radius = anchor_width / 2.0
            if system_type == "Taut":
              body_cx = pad_x
              body_cy = pad_y
            else:
              body_cx = pad_x + radius * bux
              body_cy = pad_y + radius * buy
            for th_circle in np.linspace(0.0, 2.0*np.pi, 24, endpoint=False):
              export_rows.append(create_row(
                  "Anchor_Body", primary_anc, f"Turbine_{i+1}",
                  float(a["heading_deg"]),
                  body_cx + radius*np.cos(th_circle),
                  body_cy + radius*np.sin(th_circle),
                  body_z_top, "#800080",
                  f"Inverted Bridle Anchor Body {b_idx} ({primary_anc})"
              ))

          # Branch-specific TDP/DDP: use the actual exported branch XY/Z curve,
          # not a centreline approximation between branches.
          bp = next((v for v in branch_profiles if int(v.get("branch", -1)) == b_idx), None)
          if bp is not None:
            _branch_sub_x_raw = _bg_data.get("sub_x")
            _branch_sub_x = np.asarray(
                _branch_sub_x_raw if _branch_sub_x_raw is not None else [],
                dtype=float,
            )
            try:
              _branch_x_td = float(_bg_data.get("X_td"))
            except (TypeError, ValueError):
              _branch_x_td = np.nan
            _bx3 = np.asarray(bp.get("x", []), dtype=float)
            _by3 = np.asarray(bp.get("y", []), dtype=float)
            _bz3 = np.asarray(bp.get("z", []), dtype=float)
            _pfx = np.asarray(bp.get("profile_x", []), dtype=float)
            if (_bx3.size >= 2 and _by3.size == _bx3.size and _bz3.size == _bx3.size and
                _pfx.size == _bx3.size):
              for marker_name, xq, fallback_index in (
                  ("DDP", float(_branch_sub_x[-1]) if _branch_sub_x.size else None, 0),
                  ("TDP", float(_branch_x_td) if np.isfinite(_branch_x_td) else None, -1),
              ):
                if xq is None:
                  continue
                if np.nanmin(_pfx) <= xq <= np.nanmax(_pfx):
                  xx = float(np.interp(xq, _pfx, _bx3))
                  yy = float(np.interp(xq, _pfx, _by3))
                  zz = float(np.interp(xq, _pfx, _bz3))
                else:
                  xx = float(_bx3[fallback_index]); yy = float(_by3[fallback_index]); zz = float(_bz3[fallback_index])
                export_rows.append(create_row(
                    "Anchor_Component", marker_name, f"Turbine_{i+1}",
                    float(a["heading_deg"]), xx, yy, zz, "#FFFF00",
                    f"{marker_name} for Inverted Bridle Branch T{i+1}L{line_idx}{'a' if b_idx == 1 else 'b'}"
                ))

        # Ordered physical branch geometry.
        for bp in branch_profiles:
          bidx = int(bp.get("branch", 0))
          suffix = "a" if bidx == 1 else "b"
          branch_name = f"T{i+1}L{line_idx}{suffix}"
          bx3 = np.asarray(bp.get("x", []), dtype=float)
          by3 = np.asarray(bp.get("y", []), dtype=float)
          bz3 = np.asarray(bp.get("z", []), dtype=float)
          if bx3.size < 2 or by3.size != bx3.size or bz3.size != bx3.size:
            continue
          for px, py, pz in zip(bx3, by3, bz3):
            export_rows.append(create_row(
                "Mooring_Line", branch_name, f"Turbine_{i+1}",
                float(angle), float(px), float(py), float(pz),
                "#00BFFF" if system_type == "Taut" else "#000000",
                f"Inverted Bridle Branch {branch_name}"
            ))

  return export_rows


# =========================================================================
# UNIFIED DASHBOARD GUI WINDOW
# =========================================================================

# =========================================================================
# RASTER-DEFINED SEABED / TERRAIN-AWARE DASHBOARD HELPERS
# =========================================================================
class _RasterSampler:
  """Reuse one raster handle and CRS transformer for a group of samples."""

  def __init__(self, raster_path):
    self.raster_path = raster_path
    self._src = None
    self._transformer = None
    self._open_attempted = False

  def _ensure_open(self):
    if self._open_attempted:
      return self._src is not None
    self._open_attempted = True
    try:
      _ensure_geo_dependencies()
      self._src = rasterio.open(self.raster_path)
      raster_crs = self._src.crs or "EPSG:4326"
      self._transformer = Transformer.from_crs(
          "EPSG:3857", raster_crs, always_xy=True
      )
      return True
    except Exception as exc:
      self.close()
      print(f"[WARNING] Raster seabed sampling failed: {exc}")
      return False

  def sample(self, x_epsg3857, y_epsg3857):
    """Return raster Z/elevation at EPSG:3857 locations."""
    x_arr = np.asarray(x_epsg3857, dtype=float)
    y_arr = np.asarray(y_epsg3857, dtype=float)
    scalar = x_arr.ndim == 0
    x_flat = x_arr.reshape(-1)
    y_flat = y_arr.reshape(-1)
    vals = np.full_like(x_flat, np.nan, dtype=float)
    if not self._ensure_open():
      return float(vals[0]) if scalar else vals.reshape(x_arr.shape)
    try:
      rx, ry = self._transformer.transform(x_flat, y_flat)
      vals = np.asarray([
          float(value[0]) if np.isfinite(value[0]) else np.nan
          for value in self._src.sample(zip(rx, ry), indexes=1)
      ], dtype=float)
    except Exception as exc:
      print(f"[WARNING] Raster seabed sampling failed: {exc}")
    return float(vals[0]) if scalar else vals.reshape(x_arr.shape)

  def close(self):
    if self._src is not None:
      try:
        self._src.close()
      except Exception:
        pass
    self._src = None
    self._transformer = None

  def __del__(self):
    # Covers exceptional exits from an interactive render before its explicit
    # close call is reached.
    self.close()


def sample_raster_seabed(raster_path, x_epsg3857, y_epsg3857):
  """Return raster Z/elevation at EPSG:3857 locations."""
  sampler = _RasterSampler(raster_path)
  try:
    return sampler.sample(x_epsg3857, y_epsg3857)
  finally:
    sampler.close()


# =========================================================================
# STANDALONE PYVISTA SEISMIC VOLUME VIEWER
# =========================================================================
# These are deliberately the same fields written by ``export_system_to_csv``.
# The viewer consumes them in memory, rather than saving/re-reading a CSV or
# looking for a raster inside a QGIS project.  That keeps the physical bridle,
# anchor and terrain-resolved line coordinates authoritative in one place.
_SEISMIC_EXPORT_COLUMNS = (
    "Feature_Type", "Sub_Type", "Turbine_ID", "Line_Heading_Deg",
    "X_Coord", "Y_Coord", "Z_Coord", "Max_Penetration_Depth_m",
    "Anchor_Width_m", "Anchor_Height_m", "Anchor_Length_m",
    "Cylindrical_Diameter_m", "Cylindrical_Length_m", "Max_Water_Depth_m",
    "Seabed_Level_m", "Sea_Level_m", "Local_Seabed_Z_m", "Color_Hex",
    "Description",
)


def launch_seismic_volume_calculation(
    data,
    raster_path,
    *,
    export_rows=None,
    parent_figure=None,
    grid_resolution_m=5.0,
):
  """Open the standalone, terrain-aware PyVista UHRS seismic viewer.

  ``raster_path`` is always the raster selected in the Main.py workflow.  No
  QGIS project, QGIS layer, clipped TIFF, or saved CSV is consulted here.  The
  in-memory export rows retain the exact current solver and inverted-bridle
  geometry while the regular EPSG:3857 grid replaces the former QGIS geometry
  operations for the four seismic-footprint scenarios.
  """
  try:
    import pyvista as pv
    import vtk
    import rasterio
    from matplotlib.path import Path as MatplotlibPath
    from scipy.spatial import Delaunay
    from rasterio.transform import from_bounds
    from rasterio.warp import Resampling, reproject
  except ImportError as exc:
    print(
        "[ERROR] Seismic Volume Calculation needs PyVista, rasterio and "
        f"SciPy in the Main.py Python environment: {exc}"
    )
    return False

  if not isinstance(data, dict):
    print("[ERROR] Seismic Volume Calculation received no active mooring result.")
    return False
  if not raster_path or not os.path.isfile(raster_path):
    print("[ERROR] Select a valid raster file before opening Seismic Volume Calculation.")
    return False

  try:
    _ensure_geo_dependencies()

    if export_rows is None:
      origin_x, origin_y = latlon_to_epsg3857(
          float(data["center_lat"]), float(data["center_lon"])
      )
      export_rows = generate_export_rows(data, origin_x, origin_y)
    if not export_rows:
      print("[ERROR] No mooring geometry is available for seismic-volume calculation.")
      return False

    def as_float(value, default=np.nan):
      try:
        result = float(value)
        return result if np.isfinite(result) else default
      except (TypeError, ValueError):
        return default

    def xy(record):
      return np.array((
          as_float(record.get("X_Coord")),
          as_float(record.get("Y_Coord")),
      ), dtype=float)

    def xyz(record):
      return np.array((
          as_float(record.get("X_Coord")),
          as_float(record.get("Y_Coord")),
          as_float(record.get("Z_Coord")),
      ), dtype=float)

    records = [
        dict(zip(_SEISMIC_EXPORT_COLUMNS, row))
        for row in export_rows
        if isinstance(row, (list, tuple)) and len(row) >= len(_SEISMIC_EXPORT_COLUMNS)
    ]
    if not records:
      print("[ERROR] The active mooring export geometry is incomplete.")
      return False

    perimeter_records = [
        record for record in records
        if str(record.get("Feature_Type", "")).strip() == "Farm_Perimeter"
        and np.all(np.isfinite(xy(record)))
    ]
    perimeter_xy = np.asarray([xy(record) for record in perimeter_records], dtype=float)
    if perimeter_xy.shape[0] < 3:
      geometry_xy = np.asarray([
          xy(record) for record in records if np.all(np.isfinite(xy(record)))
      ], dtype=float)
      if geometry_xy.size == 0:
        print("[ERROR] The active mooring export contains no valid XY coordinates.")
        return False
      margin = max(50.0, float(data.get("anchor_radius", 0.0)) * 0.10)
      xmin, ymin = np.min(geometry_xy, axis=0) - margin
      xmax, ymax = np.max(geometry_xy, axis=0) + margin
      perimeter_xy = np.array([
          (xmin, ymin), (xmax, ymin), (xmax, ymax), (xmin, ymax),
      ], dtype=float)

    xmin, ymin = np.min(perimeter_xy, axis=0)
    xmax, ymax = np.max(perimeter_xy, axis=0)
    if not (np.isfinite(xmin) and np.isfinite(ymin) and xmax > xmin and ymax > ymin):
      print("[ERROR] The farm perimeter has invalid dimensions.")
      return False

    # Read and reproject only the user-selected raster into a lightweight,
    # regular EPSG:3857 grid.  This is much faster than per-node opening or
    # sampling and does not create a QGIS-derived intermediate raster file.
    target = max(2.0, float(grid_resolution_m))
    # Five-metre target cells give ten samples across the 50 m acquisition
    # corridor where the farm extent permits.  The cap keeps wide-farm mesh
    # generation responsive while SDF interpolation removes raster-cell steps.
    min_nodes, max_nodes = 12, 257
    nx = int(np.clip(math.ceil((xmax - xmin) / target) + 1, min_nodes, max_nodes))
    ny = int(np.clip(math.ceil((ymax - ymin) / target) + 1, min_nodes, max_nodes))
    dx = (xmax - xmin) / max(nx - 1, 1)
    dy = (ymax - ymin) / max(ny - 1, 1)
    dst_transform = from_bounds(
        xmin - dx * 0.5, ymin - dy * 0.5,
        xmax + dx * 0.5, ymax + dy * 0.5,
        nx, ny,
    )
    top_z = np.full((ny, nx), np.nan, dtype=np.float32)
    with rasterio.open(raster_path) as raster_source:
      reproject(
          source=rasterio.band(raster_source, 1),
          destination=top_z,
          src_transform=raster_source.transform,
          src_crs=raster_source.crs or "EPSG:3857",
          src_nodata=raster_source.nodata,
          dst_transform=dst_transform,
          dst_crs="EPSG:3857",
          dst_nodata=np.nan,
          resampling=Resampling.bilinear,
      )
    valid_z = np.isfinite(top_z)
    if not np.any(valid_z):
      print("[ERROR] The selected raster has no valid elevation data over this farm perimeter.")
      return False

    # Main.py treats bathymetry as elevations below sea level.  Retain the
    # QGIS viewer's compatibility for legacy positive-depth rasters without
    # altering the original user file.
    if float(np.nanmedian(top_z[valid_z])) > 0.0:
      top_z[valid_z] = -np.abs(top_z[valid_z])

    x_nodes = np.linspace(xmin, xmax, nx, dtype=float)
    y_nodes = np.linspace(ymax, ymin, ny, dtype=float)
    X, Y = np.meshgrid(x_nodes, y_nodes)
    cell_x = 0.25 * (X[:-1, :-1] + X[:-1, 1:] + X[1:, :-1] + X[1:, 1:])
    cell_y = 0.25 * (Y[:-1, :-1] + Y[:-1, 1:] + Y[1:, :-1] + Y[1:, 1:])
    cell_top_z = 0.25 * (
        top_z[:-1, :-1] + top_z[:-1, 1:] + top_z[1:, :-1] + top_z[1:, 1:]
    )
    cell_valid = (
        valid_z[:-1, :-1] & valid_z[:-1, 1:]
        & valid_z[1:, :-1] & valid_z[1:, 1:]
    )
    cell_points = np.column_stack((cell_x.ravel(), cell_y.ravel()))

    def polygon_mask(vertices):
      vertices = np.asarray(vertices, dtype=float)
      if vertices.ndim != 2 or vertices.shape[0] < 3:
        return np.zeros(cell_x.shape, dtype=bool)
      # ``contains_points`` closes an open path itself. Passing ``closed=True``
      # would turn the final supplied vertex into CLOSEPOLY metadata and omit
      # it, which incorrectly reduced a rectangular perimeter to a triangle.
      return MatplotlibPath(vertices).contains_points(
          cell_points, radius=1.0e-7
      ).reshape(cell_x.shape)

    farm_mask = polygon_mask(perimeter_xy) & cell_valid
    if not np.any(farm_mask):
      print("[ERROR] No valid user-raster cells fall inside the farm perimeter.")
      return False

    turbines = {}
    padeyes_by_turbine = {}
    # The physical export can format a heading as either ``0`` or ``0.0``.
    # Use a normalised key only for component-to-line matching so the
    # optimisation footprint always finds the matching padeye.
    padeyes_by_turbine_heading = {}
    anchor_bodies_by_turbine_heading = {}
    line_groups = {}
    anchor_body_groups = {}
    point_groups = {"Padeye": [], "TDP": [], "DDP": [], "Fairlead": [], "Joints": []}

    def heading_index_key(value):
      text = str(value if value is not None else "").strip()
      try:
        return f"{float(text) % 360.0:.9g}"
      except (TypeError, ValueError):
        return text

    for record in records:
      feature_type = str(record.get("Feature_Type", "")).strip()
      subtype = str(record.get("Sub_Type", "")).strip()
      turbine_id = str(record.get("Turbine_ID", "")).strip()
      point = xyz(record)
      if feature_type == "Turbine" and np.all(np.isfinite(point)):
        turbines[turbine_id] = point
      elif feature_type == "Mooring_Line" and np.all(np.isfinite(point)):
        key = (turbine_id, str(record.get("Line_Heading_Deg", "")).strip(), subtype)
        group = line_groups.setdefault(key, {"points": [], "color": "#1f2937"})
        group["points"].append(point)
        color = str(record.get("Color_Hex", "")).strip()
        if color:
          group["color"] = color
      elif feature_type == "Anchor_Body" and np.all(np.isfinite(point)):
        heading_key = heading_index_key(record.get("Line_Heading_Deg", ""))
        anchor_bodies_by_turbine_heading.setdefault(
            (turbine_id, heading_key), []
        ).append(point)
        key = (
            turbine_id, str(record.get("Line_Heading_Deg", "")).strip(),
            subtype, str(record.get("Description", "")).strip(),
        )
        group = anchor_body_groups.setdefault(key, {"points": [], "record": record})
        group["points"].append(point)
      elif feature_type == "Anchor_Component" and np.all(np.isfinite(point)):
        if subtype == "Padeye":
          padeyes_by_turbine.setdefault(turbine_id, []).append(point[:2])
          heading_key = heading_index_key(record.get("Line_Heading_Deg", ""))
          padeyes_by_turbine_heading.setdefault(
              (turbine_id, heading_key), []
          ).append(point)
          point_groups["Padeye"].append(point)
        elif subtype in ("TDP", "DDP", "Fairlead"):
          point_groups[subtype].append(point)
      elif feature_type == "Mooring_Joint" and np.all(np.isfinite(point)):
        point_groups["Joints"].append(point)

    if not turbines:
      print("[ERROR] The active export has no turbine locations.")
      return False

    def decimate_polyline(points, maximum_segments=32):
      points = np.asarray(points, dtype=float)
      if points.shape[0] < 2:
        return points
      keep = np.ones(points.shape[0], dtype=bool)
      keep[1:] = np.linalg.norm(np.diff(points, axis=0), axis=1) > 1.0e-7
      points = points[keep]
      if points.shape[0] <= maximum_segments + 1:
        return points
      indices = np.unique(np.linspace(0, points.shape[0] - 1, maximum_segments + 1).astype(int))
      return points[indices]

    def extend_anchor_end(points, turbine_id, extension_m=50.0):
      """Extend only a physical padeye end, matching the former corridor rule."""
      points = np.asarray(points, dtype=float).copy()
      if points.shape[0] < 2:
        return points
      candidates = np.asarray(padeyes_by_turbine.get(turbine_id, []), dtype=float)
      if candidates.ndim != 2 or candidates.shape[0] == 0:
        return points
      end_distances = [
          float(np.min(np.linalg.norm(candidates - points[index, :2], axis=1)))
          for index in (0, -1)
      ]
      endpoint_index = 0 if end_distances[0] <= end_distances[1] else -1
      if end_distances[0 if endpoint_index == 0 else 1] > max(5.0, target * 1.5):
        return points
      if endpoint_index == 0:
        direction = points[0, :2] - points[1, :2]
      else:
        direction = points[-1, :2] - points[-2, :2]
      length = float(np.linalg.norm(direction))
      if length <= 1.0e-8:
        return points
      extension = points[endpoint_index].copy()
      extension[:2] += direction / length * float(extension_m)
      return np.vstack((extension, points)) if endpoint_index == 0 else np.vstack((points, extension))

    def canonical_regular_route(points, turbine_id, heading, extension_m=50.0):
      """Recreate the original QGIS route used for a normal-system corridor.

      The physical CSV profile starts/ends at its individual fairlead.  Those
      fairleads are intentionally offset around the turbine, so using the raw
      profile as a survey route leaves otherwise connected swaths with no
      common endpoint.  The former working QGIS viewer instead used the
      turbine centre, the matching Padeye (or anchor-body fallback), and a
      50 m radial extension.  Keep that construction for the *optimisation*
      footprint only; raw profiles remain available for physical-line display.
      """
      points = np.asarray(points, dtype=float)
      turbine_point = np.asarray(turbines.get(turbine_id, ()), dtype=float)
      if (
          points.ndim != 2 or points.shape[0] < 2
          or turbine_point.shape[0] < 3
          or not np.all(np.isfinite(turbine_point))
      ):
        return points

      planar_distances = np.linalg.norm(points[:, :2] - turbine_point[:2], axis=1)
      raw_anchor = points[int(np.nanargmax(planar_distances))].copy()
      component_key = (turbine_id, heading_index_key(heading))
      padeyes = np.asarray(
          padeyes_by_turbine_heading.get(component_key, []), dtype=float,
      )
      if padeyes.ndim == 1 and padeyes.size:
        padeyes = padeyes.reshape(1, -1)

      if padeyes.ndim == 2 and padeyes.shape[0] and padeyes.shape[1] >= 3:
        closest = int(np.argmin(np.linalg.norm(
            padeyes[:, :2] - raw_anchor[:2], axis=1
        )))
        anchor_point = padeyes[closest, :3].copy()
      else:
        anchor_body = np.asarray(
            anchor_bodies_by_turbine_heading.get(component_key, []), dtype=float,
        )
        if anchor_body.ndim == 1 and anchor_body.size:
          anchor_body = anchor_body.reshape(1, -1)
        if (
            anchor_body.ndim == 2 and anchor_body.shape[0]
            and anchor_body.shape[1] >= 3
        ):
          anchor_point = np.mean(anchor_body[:, :3], axis=0)
        else:
          # Old exports without explicit components still get a stable,
          # turbine-centred route from their farthest physical profile point.
          anchor_point = raw_anchor[:3].copy()

      direction = anchor_point[:2] - turbine_point[:2]
      length = float(np.linalg.norm(direction))
      if length <= 1.0e-8:
        return np.vstack((turbine_point, anchor_point))
      extension = anchor_point.copy()
      extension[:2] += direction / length * float(extension_m)
      return np.vstack((turbine_point, extension))

    def polyline_buffer_mask(points, width_m, x_grid=cell_x, y_grid=cell_y):
      """Return a round-capped route-buffer mask on a supplied regular grid."""
      points = np.asarray(points, dtype=float)
      if points.shape[0] < 2:
        return np.zeros(np.asarray(x_grid).shape, dtype=bool)
      x_grid = np.asarray(x_grid, dtype=float)
      y_grid = np.asarray(y_grid, dtype=float)
      result = np.zeros(x_grid.shape, dtype=bool)
      half_width_sq = (float(width_m) * 0.5) ** 2
      for p0, p1 in zip(points[:-1, :2], points[1:, :2]):
        direction = p1 - p0
        length_sq = float(np.dot(direction, direction))
        if length_sq <= 1.0e-12:
          distance_sq = (x_grid - p0[0]) ** 2 + (y_grid - p0[1]) ** 2
        else:
          factor = np.clip(
              ((x_grid - p0[0]) * direction[0] + (y_grid - p0[1]) * direction[1]) / length_sq,
              0.0, 1.0,
          )
          closest_x = p0[0] + factor * direction[0]
          closest_y = p0[1] + factor * direction[1]
          distance_sq = (x_grid - closest_x) ** 2 + (y_grid - closest_y) ** 2
        result |= distance_sq <= half_width_sq
      return result

    def bridle_branch_from_subtype(value):
      """Return (profile, branch) for an exported inverted-bridle branch."""
      match = re.match(r"^(T\d+L\d+)([ab])$", str(value).strip(), flags=re.IGNORECASE)
      if match is None:
        return None, None
      return match.group(1).upper(), match.group(2).lower()

    # The original QGIS volume tool deliberately represented an inverted
    # bridle as a turbine-to-common-joint trunk and two straight joint-to-
    # anchor arms.  Reusing the sampled branch points here made bends in the
    # corridor footprint and, more importantly, left the two sibling arms
    # disconnected when their exported lower-joint coordinates differed by a
    # tiny numerical amount.  Restore that canonical route construction for
    # the optimisation footprint; the raw sampled geometry is still rendered
    # below as the physical mooring line.
    bridle_records = {}
    for line_key, group in line_groups.items():
      turbine_id, _heading, subtype = line_key
      profile_id, branch = bridle_branch_from_subtype(subtype)
      if turbine_id in turbines and profile_id is not None:
        bridle_records.setdefault((turbine_id, profile_id), {})[branch] = (line_key, group)

    route_paths = []
    processed_bridle_keys = set()
    canonical_bridle_turbines = set()
    for (turbine_id, _profile_id), branches in sorted(bridle_records.items()):
      branch_data = []
      turbine_xy = np.asarray(turbines[turbine_id][:2], dtype=float)
      for branch_label in ("a", "b"):
        record = branches.get(branch_label)
        if record is None:
          continue
        line_key, group = record
        path = decimate_polyline(group["points"])
        if path.shape[0] < 2:
          continue
        first, last = path[0].copy(), path[-1].copy()
        first_distance = float(np.linalg.norm(first[:2] - turbine_xy))
        last_distance = float(np.linalg.norm(last[:2] - turbine_xy))
        anchor, joint = (first, last) if first_distance >= last_distance else (last, first)
        if float(np.linalg.norm(anchor[:2] - joint[:2])) <= 1.0e-8:
          continue
        branch_data.append((line_key, anchor, joint, group["color"]))

      # An incomplete bridle falls through to the normal route construction
      # below rather than inventing a common trunk from one branch.  A
      # complete pair reproduces the prior working QGIS construction exactly.
      if len(branch_data) != 2:
        continue
      canonical_bridle_turbines.add(turbine_id)
      common_joint = np.mean([joint for _key, _anchor, joint, _color in branch_data], axis=0)
      turbine_point = np.asarray(turbines[turbine_id], dtype=float)
      trunk = np.vstack((turbine_point, common_joint))
      if float(np.linalg.norm(np.diff(trunk[:, :2], axis=0)[0])) > 1.0e-8:
        route_paths.append((turbine_id, trunk, branch_data[0][3]))
      for line_key, anchor, _joint, color in branch_data:
        direction = anchor[:2] - common_joint[:2]
        length = float(np.linalg.norm(direction))
        if length <= 1.0e-8:
          continue
        extension = anchor.copy()
        extension[:2] += direction / length * 50.0
        route_paths.append((
            turbine_id,
            np.vstack((common_joint, anchor, extension)),
            color,
        ))
        processed_bridle_keys.add(line_key)

    # For regular systems, reconstruct the original QGIS survey route from
    # the turbine centre to the matching padeye and then extend it by 50 m.
    # The export for a bridle also contains a sampled generic trunk (for
    # physical line display), but it must not be added to the optimisation
    # footprint: the canonical turbine-to-common-joint trunk above is its
    # original QGIS equivalent.
    for line_key, group in line_groups.items():
      turbine_id, heading, _subtype = line_key
      if (
          turbine_id not in turbines
          or turbine_id in canonical_bridle_turbines
          or line_key in processed_bridle_keys
      ):
        continue
      path = decimate_polyline(group["points"])
      if path.shape[0] < 2:
        continue
      route = canonical_regular_route(path, turbine_id, heading)
      if route.shape[0] >= 2:
        route_paths.append((turbine_id, route, group["color"]))

    containment_points = {turbine_id: [] for turbine_id in turbines}
    line_angles = {turbine_id: [] for turbine_id in turbines}
    corridor_mask = np.zeros(cell_x.shape, dtype=bool)
    intermediate_buffer_m = 50.0
    for turbine_id, route, _color in route_paths:
      corridor_mask |= polyline_buffer_mask(route, 50.0)
      stride = max(1, int(math.ceil(route.shape[0] / 24.0)))
      samples = route[::stride, :2]
      if not np.allclose(samples[-1], route[-1, :2]):
        samples = np.vstack((samples, route[-1, :2]))
      for sample in samples:
        containment_points[turbine_id].append(sample)
        circle = np.linspace(0.0, 2.0 * np.pi, 8, endpoint=False)
        containment_points[turbine_id].extend(
            sample + intermediate_buffer_m * np.column_stack((np.cos(circle), np.sin(circle)))
        )
      turbine_xy = turbines[turbine_id][:2]
      end_index = int(np.argmax(np.linalg.norm(route[:, :2] - turbine_xy, axis=1)))
      endpoint = route[end_index, :2]
      vector = endpoint - turbine_xy
      if float(np.linalg.norm(vector)) > 1.0e-8:
        line_angles[turbine_id].append(float(np.degrees(np.arctan2(vector[1], vector[0])) % 360.0))

    # The optimised-corridor scenario is deliberately just the strict union
    # of 25 m buffers around its survey routes.  Do not fill internal voids
    # or add connector wedges: overlapping swaths remain a natural overlap,
    # while genuinely separate corridors stay separate.
    corridor_mask &= farm_mask

    def triangle_vertices(center, side, angle_rad):
      angles = np.radians((90.0, 210.0, 330.0)) + float(angle_rad)
      radius = float(side) / math.sqrt(3.0)
      return np.column_stack((
          center[0] + radius * np.cos(angles),
          center[1] + radius * np.sin(angles),
      ))

    def triangle_orientation(angles):
      if not angles:
        return math.radians(15.0)
      best_angle, best_score = math.radians(15.0), float("inf")
      for candidate_deg in np.arange(0.0, 120.0, 0.5):
        vertices = (np.array((90.0, 210.0, 330.0)) + candidate_deg) % 360.0
        score = sum(min(abs((angle - vertex + 180.0) % 360.0 - 180.0) for vertex in vertices) ** 2
                    for angle in angles)
        if score < best_score:
          best_angle, best_score = math.radians(float(candidate_deg)), float(score)
      return best_angle

    def polygon_contains_all(vertices, points):
      if not points:
        return True
      return bool(np.all(MatplotlibPath(vertices).contains_points(
          np.asarray(points, dtype=float), radius=1.0e-7
      )))

    def minimum_triangle_side(center, points, angle):
      if not points:
        return 1.0
      radius = max(float(np.linalg.norm(np.asarray(point) - center)) for point in points)
      low, high = 1.0, max(2.0, radius * 2.0 + 2.0 * intermediate_buffer_m)
      while high < 1.0e6 and not polygon_contains_all(triangle_vertices(center, high, angle), points):
        high *= 2.0
      for _ in range(42):
        middle = (low + high) * 0.5
        if polygon_contains_all(triangle_vertices(center, middle, angle), points):
          high = middle
        else:
          low = middle
      return high

    square_mask = np.zeros(cell_x.shape, dtype=bool)
    triangle_mask = np.zeros(cell_x.shape, dtype=bool)
    square_polygons = []
    triangle_polygons = []
    for turbine_id, turbine_xyz in turbines.items():
      centre = turbine_xyz[:2]
      points = containment_points.get(turbine_id, []) or [centre]
      relative = np.asarray(points, dtype=float) - centre
      square_side = max(1.0, 2.0 * float(np.max(np.abs(relative))))
      orientation = triangle_orientation(line_angles.get(turbine_id, []))
      triangle_side = minimum_triangle_side(centre, points, orientation)
      common_side = max(square_side, triangle_side)
      square_vertices = np.array([
          (centre[0] - common_side * 0.5, centre[1] - common_side * 0.5),
          (centre[0] + common_side * 0.5, centre[1] - common_side * 0.5),
          (centre[0] + common_side * 0.5, centre[1] + common_side * 0.5),
          (centre[0] - common_side * 0.5, centre[1] + common_side * 0.5),
      ], dtype=float)
      triangle_shape = triangle_vertices(centre, common_side, orientation)
      square_polygons.append(square_vertices)
      triangle_polygons.append(triangle_shape)
      square_mask |= polygon_mask(square_vertices)
      triangle_mask |= polygon_mask(triangle_shape)
    square_mask &= farm_mask
    triangle_mask &= farm_mask

    penetration_values = [
        as_float(record.get("Max_Penetration_Depth_m"), 0.0)
        for record in records
        if str(record.get("Feature_Type", "")).strip() == "Anchor_Body"
    ]
    base_penetration = max(
        [value for value in penetration_values if np.isfinite(value) and value > 0.0]
        or [as_float(data.get("anchor_depth"), 7.5)]
    )
    target_depth = max(0.1, 2.0 * float(base_penetration))
    bottom_z = float(np.nanmin(top_z[valid_z]) - target_depth)
    cell_area_m2 = abs(float(dx * dy))

    def make_volume_mesh(mask):
      mask = np.asarray(mask, dtype=bool) & cell_valid
      row_indices, col_indices = np.nonzero(mask)
      if row_indices.size == 0:
        return None
      node_count = nx * ny
      safe_top = np.where(np.isfinite(top_z), top_z, bottom_z)
      base_points = np.column_stack((
          X.ravel(order="F"), Y.ravel(order="F"),
          np.full(node_count, bottom_z, dtype=float),
      ))
      top_points = np.column_stack((
          X.ravel(order="F"), Y.ravel(order="F"), safe_top.ravel(order="F"),
      ))
      b0 = row_indices + col_indices * ny
      b1 = row_indices + (col_indices + 1) * ny
      b2 = (row_indices + 1) + (col_indices + 1) * ny
      b3 = (row_indices + 1) + col_indices * ny
      cells = np.column_stack((
          np.full(row_indices.size, 8, dtype=np.int64),
          b0, b1, b2, b3,
          b0 + node_count, b1 + node_count, b2 + node_count, b3 + node_count,
      )).ravel()
      types = np.full(row_indices.size, pv.CellType.HEXAHEDRON, dtype=np.uint8)
      return pv.UnstructuredGrid(cells, types, np.vstack((base_points, top_points)))

    def sample_volume_top(xy_points):
      """Bilinearly sample the display terrain on the regular volume grid."""
      points = np.asarray(xy_points, dtype=float).reshape(-1, 2)
      safe_top = np.where(np.isfinite(top_z), top_z, bottom_z).astype(float)
      columns = np.clip((points[:, 0] - xmin) / max(dx, 1.0e-12), 0.0, nx - 1.0)
      rows = np.clip((ymax - points[:, 1]) / max(abs(dy), 1.0e-12), 0.0, ny - 1.0)
      col0 = np.floor(columns).astype(int)
      row0 = np.floor(rows).astype(int)
      col1 = np.minimum(col0 + 1, nx - 1)
      row1 = np.minimum(row0 + 1, ny - 1)
      tx = columns - col0
      ty = rows - row0
      return (
          safe_top[row0, col0] * (1.0 - tx) * (1.0 - ty)
          + safe_top[row0, col1] * tx * (1.0 - ty)
          + safe_top[row1, col0] * (1.0 - tx) * ty
          + safe_top[row1, col1] * tx * ty
      )

    def capsule_outline(start, end, radius=25.0, arc_segments=8):
      """Return a round-capped 2-D strip matching QGIS ``buffer(..., 8)``."""
      start, end = np.asarray(start, dtype=float), np.asarray(end, dtype=float)
      direction = end - start
      length = float(np.linalg.norm(direction))
      if length <= 1.0e-7:
        return None
      heading = math.atan2(direction[1], direction[0])
      # Start: left-to-right around the outward (rear) end; end: right-to-left
      # around the forward end.  Together these form a CCW capsule ring.
      start_angles = np.linspace(heading + math.pi * 0.5, heading + math.pi * 1.5, arc_segments + 1)
      end_angles = np.linspace(heading - math.pi * 0.5, heading + math.pi * 0.5, arc_segments + 1)
      ring = np.vstack((
          start + float(radius) * np.column_stack((np.cos(start_angles), np.sin(start_angles))),
          end + float(radius) * np.column_stack((np.cos(end_angles), np.sin(end_angles))),
      ))
      signed_area = 0.5 * float(np.sum(
          ring[:, 0] * np.roll(ring[:, 1], -1)
          - ring[:, 1] * np.roll(ring[:, 0], -1)
      ))
      return ring if signed_area >= 0.0 else ring[::-1]

    def capsule_volume_mesh(start, end, width_m=50.0):
      """Build one smooth, terrain-following route-volume shell.

      This is the Main.py equivalent of the original QGIS round-capped route
      buffer.  It is display-only: the raster-cell union remains the source
      of the reproducible survey-area and volume metrics.
      """
      ring = capsule_outline(start, end, radius=float(width_m) * 0.5, arc_segments=8)
      if ring is None:
        return None
      centre = np.mean(ring, axis=0, keepdims=True)
      top_xy = np.vstack((centre, ring))
      top_points = np.column_stack((top_xy, sample_volume_top(top_xy)))
      point_count = top_points.shape[0]
      bottom_points = top_points.copy()
      bottom_points[:, 2] = bottom_z
      faces = []
      for index in range(1, point_count):
        next_index = 1 if index == point_count - 1 else index + 1
        faces.extend((3, 0, index, next_index))
        faces.extend((3, point_count, point_count + next_index, point_count + index))
        faces.extend((4, index, point_count + index, point_count + next_index, next_index))
      return pv.PolyData(
          np.vstack((top_points, bottom_points)),
          faces=np.asarray(faces, dtype=np.int64),
      )

    def route_display_segments(points):
      """Collapse sampled straight mooring paths into their physical XY runs."""
      points = np.asarray(points, dtype=float)
      if points.ndim != 2 or points.shape[0] < 2:
        return []
      xy_points = points[:, :2]
      keep = [0]
      for index in range(1, len(xy_points) - 1):
        previous = xy_points[keep[-1]]
        current = xy_points[index]
        following = xy_points[index + 1]
        left, right = current - previous, following - current
        left_length, right_length = float(np.linalg.norm(left)), float(np.linalg.norm(right))
        if left_length <= 1.0e-7 or right_length <= 1.0e-7:
          continue
        # Solver samples on a straight plan-view line must remain one single
        # capsule; retaining every sample was the source of the stepped red
        # cell footprint in the migrated renderer.
        cross_z = float(left[0] * right[1] - left[1] * right[0])
        if abs(cross_z) / (left_length * right_length) > 1.0e-4:
          keep.append(index)
      keep.append(len(xy_points) - 1)
      return [
          (xy_points[first], xy_points[second])
          for first, second in zip(keep[:-1], keep[1:])
          if float(np.linalg.norm(xy_points[second] - xy_points[first])) > 1.0e-7
      ]

    def make_smooth_corridor_display_meshes():
      """Make the original-style dissolved, round-boundary display shell.

      V4.0 created an analytic QGIS buffer union, then clipped a terrain cap
      to that union.  QGIS is intentionally absent from this workflow, so the
      equivalent is made with VTK implicit polygon loops: each 50 m capsule is
      one analytic route buffer, their union is intersected with the farm
      fence, and the resulting cap is closed to the fixed volume datum.
      """
      rings = []
      meshes = []
      for _turbine_id, route, _color in route_paths:
        for start, end in route_display_segments(route):
          ring = capsule_outline(start, end, radius=25.0, arc_segments=8)
          if ring is not None:
            rings.append(ring)
          mesh = capsule_volume_mesh(start, end, width_m=50.0)
          if mesh is not None and mesh.n_cells:
            meshes.append(mesh)
      if not rings:
        return meshes

      try:
        def selection_loop(ring):
          vtk_points = vtk.vtkPoints()
          try:
            vtk_points.SetDataTypeToDouble()
          except Exception:
            pass
          for x_coord, y_coord in np.asarray(ring, dtype=float):
            vtk_points.InsertNextPoint(float(x_coord), float(y_coord), 0.0)
          loop = vtk.vtkImplicitSelectionLoop()
          loop.SetLoop(vtk_points)
          try:
            loop.SetNormal(0.0, 0.0, 1.0)
          except Exception:
            pass
          return loop

        corridor_union = vtk.vtkImplicitBoolean()
        corridor_union.SetOperationTypeToUnion()
        for ring in rings:
          corridor_union.AddFunction(selection_loop(ring))
        region = vtk.vtkImplicitBoolean()
        region.SetOperationTypeToIntersection()
        region.AddFunction(corridor_union)
        region.AddFunction(selection_loop(perimeter_xy))

        cap_grid = pv.StructuredGrid(X, Y, np.where(np.isfinite(top_z), top_z, bottom_z))
        try:
          cap_surface = cap_grid.extract_surface(algorithm="dataset_surface").triangulate()
        except TypeError:
          cap_surface = cap_grid.extract_surface().triangulate()
        clipper = vtk.vtkClipPolyData()
        clipper.SetInputData(cap_surface)
        clipper.SetClipFunction(region)
        clipper.SetValue(0.0)
        clipper.InsideOutOn()
        clipper.GenerateClippedOutputOff()
        clipper.Update()
        clipped_cap = pv.wrap(clipper.GetOutput()).copy(deep=True).triangulate()
        raw_faces = np.asarray(clipped_cap.faces, dtype=np.int64)
        if raw_faces.size == 0:
          return meshes
        face_rows = raw_faces.reshape(-1, 4).copy()
        if not np.all(face_rows[:, 0] == 3):
          return meshes
        top_xy = np.asarray(clipped_cap.points[:, :2], dtype=float)
        top_points = np.column_stack((top_xy, sample_volume_top(top_xy)))

        # Ensure a consistently upward cap, then derive the bottom and all
        # boundary walls from its indexed triangle edges, exactly as V4.0 did.
        triangles = top_points[face_rows[:, 1:4]]
        normal_z = float(np.sum(np.cross(
            triangles[:, 1] - triangles[:, 0],
            triangles[:, 2] - triangles[:, 0],
        )[:, 2]))
        if normal_z < 0.0:
          face_rows[:, [2, 3]] = face_rows[:, [3, 2]]
        edge_occurrences = {}
        for _count, first, second, third in face_rows:
          for start, end in ((first, second), (second, third), (third, first)):
            key = (min(int(start), int(end)), max(int(start), int(end)))
            edge_occurrences.setdefault(key, []).append((int(start), int(end)))
        boundary_edges = [edges[0] for edges in edge_occurrences.values() if len(edges) == 1]
        if not boundary_edges:
          return meshes
        point_count = len(top_points)
        bottom_points = top_points.copy()
        bottom_points[:, 2] = bottom_z
        bottom_faces = face_rows.copy()
        bottom_faces[:, [2, 3]] = bottom_faces[:, [3, 2]]
        bottom_faces[:, 1:4] += point_count
        wall_faces = np.asarray([
            (4, first, first + point_count, second + point_count, second)
            for first, second in boundary_edges
        ], dtype=np.int64)
        shell = pv.PolyData(
            np.vstack((top_points, bottom_points)),
            faces=np.concatenate((face_rows.ravel(), bottom_faces.ravel(), wall_faces.ravel())),
        )
        if shell.n_cells:
          return [shell]
      except Exception as exc:
        print(f"[WARNING] Smooth corridor display shell unavailable; using individual route shells: {exc}")
      return meshes

    def polygon_union_contains(points, polygons):
      """Vectorised point-in-union test, including the polygon boundary."""
      points = np.asarray(points, dtype=float).reshape(-1, 2)
      inside = np.zeros(points.shape[0], dtype=bool)
      for polygon in polygons:
        polygon = np.asarray(polygon, dtype=float)
        if polygon.ndim == 2 and polygon.shape[0] >= 3:
          inside |= MatplotlibPath(polygon).contains_points(points, radius=1.0e-4)
      return inside

    def make_smooth_volume_mesh(boundary_rings, contains_points, label):
      """Create a closed terrain shell and its metrics from one smooth footprint.

      Boundary rings provide exact plan-view edges (round route buffers,
      farm fence, squares, or triangles).  A Delaunay cap is formed from those
      edges plus valid terrain samples, with every triangle checked against the
      same footprint.  Consequently this one mesh is both what the user sees
      and what is integrated for area and survey volume.
      """
      rings = [
          np.asarray(ring, dtype=float)
          for ring in boundary_rings
          if np.asarray(ring).ndim == 2 and np.asarray(ring).shape[0] >= 3
      ]
      if not rings:
        return None, 0.0, 0.0
      terrain_samples = np.column_stack((X.ravel(), Y.ravel()))
      terrain_samples = terrain_samples[valid_z.ravel()]
      inside_samples = contains_points(terrain_samples)
      candidate_parts = [terrain_samples[inside_samples]]
      candidate_parts.extend(rings)
      candidates = np.vstack([part for part in candidate_parts if part.size])
      finite = np.all(np.isfinite(candidates), axis=1)
      candidates = candidates[finite]
      if candidates.shape[0] < 3:
        return None, 0.0, 0.0
      # Delaunay is sensitive to duplicate ring/grid vertices.  Retain each
      # original coordinate while using rounded values only as duplicate keys.
      _, unique_indices = np.unique(np.round(candidates, 7), axis=0, return_index=True)
      candidates = candidates[np.sort(unique_indices)]
      if candidates.shape[0] < 3:
        return None, 0.0, 0.0
      try:
        # EPSG:3857 northings are several million metres.  Translate and
        # scale before Qhull so its precision/joggle tolerance is related to
        # the footprint size rather than the world-coordinate magnitude.
        qhull_origin = np.mean(candidates, axis=0)
        qhull_scale = max(float(np.max(np.ptp(candidates, axis=0))), 1.0)
        simplices = Delaunay(
            (candidates - qhull_origin) / qhull_scale, qhull_options="QJ"
        ).simplices
      except Exception as exc:
        print(f"[WARNING] Could not triangulate smooth {label} footprint: {exc}")
        return None, 0.0, 0.0
      if simplices.size == 0:
        return None, 0.0, 0.0

      triangle_xy = candidates[simplices]
      edge_midpoints = np.stack((
          0.5 * (triangle_xy[:, 0] + triangle_xy[:, 1]),
          0.5 * (triangle_xy[:, 1] + triangle_xy[:, 2]),
          0.5 * (triangle_xy[:, 2] + triangle_xy[:, 0]),
      ), axis=1)
      test_points = np.concatenate((
          triangle_xy, edge_midpoints, np.mean(triangle_xy, axis=1)[:, None, :],
      ), axis=1).reshape(-1, 2)
      accepted = np.all(contains_points(test_points).reshape(-1, 7), axis=1)
      simplices = simplices[accepted]
      if simplices.size == 0:
        print(f"[WARNING] Smooth {label} footprint contains no valid triangles.")
        return None, 0.0, 0.0

      used_indices, inverse = np.unique(simplices.ravel(), return_inverse=True)
      top_xy = candidates[used_indices]
      simplices = inverse.reshape(-1, 3)
      signed_area = (
          (top_xy[simplices[:, 1], 0] - top_xy[simplices[:, 0], 0])
          * (top_xy[simplices[:, 2], 1] - top_xy[simplices[:, 0], 1])
          - (top_xy[simplices[:, 1], 1] - top_xy[simplices[:, 0], 1])
          * (top_xy[simplices[:, 2], 0] - top_xy[simplices[:, 0], 0])
      )
      reversed_faces = signed_area < 0.0
      simplices[reversed_faces, [1, 2]] = simplices[reversed_faces, [2, 1]]
      triangle_area = 0.5 * np.abs(signed_area)
      top_points = np.column_stack((top_xy, sample_volume_top(top_xy)))
      terrain_volume = float(np.sum(
          triangle_area * (np.mean(top_points[simplices, 2], axis=1) - bottom_z)
      ))
      plan_area = float(np.sum(triangle_area))

      top_faces = np.column_stack((
          np.full(simplices.shape[0], 3, dtype=np.int64), simplices,
      ))
      edge_occurrences = {}
      for first, second, third in simplices:
        for start, end in ((first, second), (second, third), (third, first)):
          key = (min(int(start), int(end)), max(int(start), int(end)))
          edge_occurrences.setdefault(key, []).append((int(start), int(end)))
      boundary_edges = [edges[0] for edges in edge_occurrences.values() if len(edges) == 1]
      if not boundary_edges:
        print(f"[WARNING] Smooth {label} footprint has no boundary shell.")
        return None, 0.0, 0.0
      point_count = top_points.shape[0]
      bottom_points = top_points.copy()
      bottom_points[:, 2] = bottom_z
      bottom_faces = top_faces.copy()
      bottom_faces[:, [2, 3]] = bottom_faces[:, [3, 2]]
      bottom_faces[:, 1:4] += point_count
      wall_faces = np.asarray([
          (4, first, first + point_count, second + point_count, second)
          for first, second in boundary_edges
      ], dtype=np.int64)
      mesh = pv.PolyData(
          np.vstack((top_points, bottom_points)),
          faces=np.concatenate((top_faces.ravel(), bottom_faces.ravel(), wall_faces.ravel())),
      )
      try:
        mesh.field_data["SmoothPlanAreaM2"] = np.asarray([plan_area])
        mesh.field_data["SmoothTerrainVolumeM3"] = np.asarray([terrain_volume])
      except Exception:
        pass
      return mesh, plan_area, terrain_volume

    def make_smooth_scenario_meshes():
      """Build and measure all four scenarios from continuous SDF footprints.

      A signed-distance field gives the exact union of round-capped route
      buffers without relying on VTK implicit-boolean nesting.  Clipping a
      triangulated terrain cap at zero interpolates the perimeter between grid
      nodes, so all scenarios have smooth edges and one shared display/metric
      topology.
      """
      def polygon_signed_distance(points, polygons):
        points = np.asarray(points, dtype=float).reshape(-1, 2)
        signed_distance = np.full(points.shape[0], np.inf, dtype=float)
        for polygon in polygons:
          polygon = np.asarray(polygon, dtype=float)
          if polygon.ndim != 2 or polygon.shape[0] < 3:
            continue
          distance = np.full(points.shape[0], np.inf, dtype=float)
          for start, end in zip(polygon, np.vstack((polygon[1:], polygon[:1]))):
            direction = end - start
            length_sq = float(np.dot(direction, direction))
            if length_sq <= 1.0e-12:
              continue
            fraction = np.clip(((points - start) @ direction) / length_sq, 0.0, 1.0)
            nearest = start + fraction[:, None] * direction
            distance = np.minimum(distance, np.linalg.norm(points - nearest, axis=1))
          inside = MatplotlibPath(polygon).contains_points(points, radius=1.0e-7)
          signed_distance = np.minimum(
              signed_distance, np.where(inside, -distance, distance)
          )
        return signed_distance

      route_segments = []
      for _turbine_id, route, _color in route_paths:
        route_segments.extend(route_display_segments(route))

      def corridor_signed_distance(points):
        points = np.asarray(points, dtype=float).reshape(-1, 2)
        signed_distance = np.full(points.shape[0], np.inf, dtype=float)
        for start, end in route_segments:
          start, end = np.asarray(start, dtype=float), np.asarray(end, dtype=float)
          direction = end - start
          length_sq = float(np.dot(direction, direction))
          if length_sq <= 1.0e-12:
            continue
          fraction = np.clip(((points - start) @ direction) / length_sq, 0.0, 1.0)
          nearest = start + fraction[:, None] * direction
          signed_distance = np.minimum(
              signed_distance, np.linalg.norm(points - nearest, axis=1) - 25.0
          )
        return signed_distance

      def repaired_corridor_signed_distance(points):
        """Strict union of the individual 25 m survey-route buffers."""
        return corridor_signed_distance(points)

      def welded_clipped_cap(cap):
        """Weld only VTK's duplicate clip vertices without cleaning faces.

        ``vtkCleanPolyData`` is tempting after a scalar clip because the
        clipper may emit the same boundary point under multiple IDs.  However,
        it can also collapse the very thin valid triangles produced at a
        diagonal round-corridor boundary into lines/vertices.  Those discarded
        cap cells showed up as the small blue notches in the gold volume.
        V4.0 fixed the equivalent issue by rebuilding fresh PolyData from
        indexed triangles; replicate that here while retaining every valid
        triangle and removing only duplicate vertices/faces.
        """
        cap = cap.triangulate()
        raw_faces = np.asarray(cap.faces, dtype=np.int64)
        if raw_faces.size == 0 or raw_faces.size % 4 != 0:
          return pv.PolyData()
        face_rows = raw_faces.reshape(-1, 4)
        if not np.all(face_rows[:, 0] == 3):
          return pv.PolyData()
        points = np.asarray(cap.points, dtype=float)
        if points.shape[0] == 0:
          return pv.PolyData()
        # Intersections on a regular 5 m grid agree much more closely than
        # this tolerance; using local XY keys avoids rounding large EPSG:3857
        # coordinates directly while preserving all physical triangles.
        xy_origin = np.mean(points[:, :2], axis=0)
        point_keys = np.rint((points[:, :2] - xy_origin) / 1.0e-7).astype(np.int64)
        _keys, unique_indices, inverse = np.unique(
            point_keys, axis=0, return_index=True, return_inverse=True,
        )
        triangle_indices = inverse[face_rows[:, 1:4]]
        non_degenerate = np.all(
            triangle_indices != np.roll(triangle_indices, 1, axis=1), axis=1
        )
        triangle_indices = triangle_indices[non_degenerate]
        if triangle_indices.size == 0:
          return pv.PolyData()
        # Removing identical overlapping triangles prevents coplanar depth
        # artefacts, but does not alter an adjacent triangle's boundary.
        sorted_indices = np.sort(triangle_indices, axis=1)
        _unique_faces, face_keep = np.unique(
            sorted_indices, axis=0, return_index=True,
        )
        triangle_indices = triangle_indices[np.sort(face_keep)]
        faces = np.column_stack((
            np.full(triangle_indices.shape[0], 3, dtype=np.int64),
            triangle_indices,
        ))
        return pv.PolyData(points[unique_indices], faces=faces.ravel())

      def closed_shell_from_cap(cap, label):
        raw_faces = np.asarray(cap.faces, dtype=np.int64)
        if raw_faces.size == 0 or raw_faces.size % 4 != 0:
          return None, 0.0, 0.0
        top_faces = raw_faces.reshape(-1, 4).copy()
        if not np.all(top_faces[:, 0] == 3):
          return None, 0.0, 0.0
        top_xy = np.asarray(cap.points[:, :2], dtype=float)
        top_points = np.column_stack((top_xy, sample_volume_top(top_xy)))
        triangles = top_points[top_faces[:, 1:4]]
        signed_area = (
            (triangles[:, 1, 0] - triangles[:, 0, 0]) * (triangles[:, 2, 1] - triangles[:, 0, 1])
            - (triangles[:, 1, 1] - triangles[:, 0, 1]) * (triangles[:, 2, 0] - triangles[:, 0, 0])
        )
        reversed_faces = signed_area < 0.0
        if np.any(reversed_faces):
          top_faces[reversed_faces] = top_faces[reversed_faces][:, [0, 1, 3, 2]]
        triangle_area = 0.5 * np.abs(signed_area)
        plan_area = float(np.sum(triangle_area))
        terrain_volume = float(np.sum(
            triangle_area * (np.mean(top_points[top_faces[:, 1:4], 2], axis=1) - bottom_z)
        ))
        edge_occurrences = {}
        for _count, first, second, third in top_faces:
          for start, end in ((first, second), (second, third), (third, first)):
            key = (min(int(start), int(end)), max(int(start), int(end)))
            edge_occurrences.setdefault(key, []).append((int(start), int(end)))
        boundary_edges = [edges[0] for edges in edge_occurrences.values() if len(edges) == 1]
        if not boundary_edges:
          print(f"[WARNING] Smooth {label} cap has no boundary edges.")
          return None, 0.0, 0.0
        point_count = top_points.shape[0]
        bottom_points = top_points.copy()
        bottom_points[:, 2] = bottom_z
        bottom_faces = top_faces.copy()
        bottom_faces[:, [2, 3]] = bottom_faces[:, [3, 2]]
        bottom_faces[:, 1:4] += point_count
        wall_faces = np.asarray([
            (4, first, first + point_count, second + point_count, second)
            for first, second in boundary_edges
        ], dtype=np.int64)
        mesh = pv.PolyData(
            np.vstack((top_points, bottom_points)),
            faces=np.concatenate((top_faces.ravel(), bottom_faces.ravel(), wall_faces.ravel())),
        )
        try:
          mesh.field_data["SmoothPlanAreaM2"] = np.asarray([plan_area])
          mesh.field_data["SmoothTerrainVolumeM3"] = np.asarray([terrain_volume])
        except Exception:
          pass
        return mesh, plan_area, terrain_volume

      def farm_signed_distance(points):
        return polygon_signed_distance(points, [perimeter_xy])

      def clipped_corridor_distance(points):
        return np.maximum(
            repaired_corridor_signed_distance(points), farm_signed_distance(points)
        )

      def clipped_triangle_distance(points):
        return np.maximum(
            polygon_signed_distance(points, triangle_polygons), farm_signed_distance(points)
        )

      def clipped_square_distance(points):
        return np.maximum(
            polygon_signed_distance(points, square_polygons), farm_signed_distance(points)
        )

      scenario_distance_functions = (
          ("Full Farm Area", farm_signed_distance),
          ("Idealised Mooring Corridors", clipped_corridor_distance),
          ("Intermediate triangular corridor", clipped_triangle_distance),
          ("Intermediate square corridor", clipped_square_distance),
      )
      terrain_surface = pv.StructuredGrid(X, Y, np.where(np.isfinite(top_z), top_z, bottom_z))
      try:
        terrain_cap_source = terrain_surface.extract_surface(
            algorithm="dataset_surface"
        ).triangulate()
      except TypeError:
        terrain_cap_source = terrain_surface.extract_surface().triangulate()
      result = {}
      cap_xy = np.asarray(terrain_cap_source.points[:, :2], dtype=float)
      for name, distance_function in scenario_distance_functions:
        try:
          cap = terrain_cap_source.copy(deep=True)
          # ``extract_surface`` changes point order, so evaluate the SDF at
          # this cap's actual XY coordinates.  Flattening the source raster
          # grid here would transpose/permutate corridors on non-square farms.
          cap.point_data["SmoothFootprintSDF"] = distance_function(cap_xy)
          # ``invert=True`` retains the negative (inside) portion and creates
          # interpolated vertices exactly on the SDF = 0 footprint boundary.
          cap = cap.clip_scalar(
              scalars="SmoothFootprintSDF", value=0.0, invert=True,
          ).triangulate()
          # VTK clipping can duplicate shared intersection points.  Rebuild a
          # welded triangle cap rather than using ``clean()``, which can erase
          # thin valid faces and leave visible notches in diagonal corridors.
          cap = welded_clipped_cap(cap)
          mesh, area_m2, volume_m3 = closed_shell_from_cap(cap, name)
        except Exception as exc:
          print(f"[WARNING] Could not build smooth {name} volume: {exc}")
          mesh, area_m2, volume_m3 = None, 0.0, 0.0
        result[name] = {
            "mesh": mesh,
            "area_m2": area_m2,
            "terrain_volume_m3": volume_m3,
        }
      return result

    def surface_mesh(mask):
      safe_top = np.where(np.isfinite(top_z), top_z, bottom_z)
      grid = pv.StructuredGrid(X, Y, safe_top)
      grid.point_data["Bathymetry"] = safe_top.ravel(order="F")
      grid.cell_data["inside"] = (np.asarray(mask, dtype=np.uint8) & cell_valid.astype(np.uint8)).ravel(order="F")
      return grid.threshold(0.5, scalars="inside", preference="cell")

    def lawnmower_segments(mask, orientation):
      mask = np.asarray(mask, dtype=bool) & cell_valid
      segments = []
      if orientation == "horizontal":
        for row_index in range(mask.shape[0]):
          columns = np.flatnonzero(mask[row_index])
          for run in np.split(columns, np.where(np.diff(columns) != 1)[0] + 1):
            if run.size:
              y_value = float(cell_y[row_index, run[0]])
              segments.append(np.array([
                  (X[row_index, run[0]], y_value),
                  (X[row_index, run[-1] + 1], y_value),
              ], dtype=float))
      else:
        for col_index in range(mask.shape[1]):
          rows = np.flatnonzero(mask[:, col_index])
          for run in np.split(rows, np.where(np.diff(rows) != 1)[0] + 1):
            if run.size:
              x_value = float(cell_x[run[0], col_index])
              segments.append(np.array([
                  (x_value, Y[run[0], col_index]),
                  (x_value, Y[run[-1] + 1, col_index]),
              ], dtype=float))
      length = float(sum(np.linalg.norm(segment[1] - segment[0]) for segment in segments))
      return segments, length

    def best_lawnmower(mask):
      horizontal, h_length = lawnmower_segments(mask, "horizontal")
      vertical, v_length = lawnmower_segments(mask, "vertical")
      return (horizontal, h_length) if h_length <= v_length else (vertical, v_length)

    sea_values = [as_float(record.get("Sea_Level_m")) for record in records]
    sea_level = next((value for value in sea_values if np.isfinite(value)), 0.0)
    procedural_seabed = float(np.nanmedian(top_z[valid_z]))
    scenario_specs = (
        ("Full Farm Area", farm_mask, "#7c3aed", 0.10),
        ("Idealised Mooring Corridors", corridor_mask, "#f59e0b", 0.30),
        ("Intermediate triangular corridor", triangle_mask, "#2563eb", 0.16),
        ("Intermediate square corridor", square_mask, "#16a34a", 0.16),
    )
    smooth_scenarios = make_smooth_scenario_meshes()
    metrics = {}
    scenario_paths = {}
    for name, mask, _color, _opacity in scenario_specs:
      active = np.asarray(mask, dtype=bool) & cell_valid
      smooth = smooth_scenarios.get(name, {})
      if smooth.get("mesh") is not None:
        # The reported area and terrain-aware volume are derived from the
        # exact same smooth shell used in the 3-D viewer.
        area_m2 = float(smooth["area_m2"])
        terrain_volume = float(smooth["terrain_volume_m3"])
      else:
        # A conservative fallback keeps exports usable if VTK/SciPy cannot
        # triangulate an invalid footprint; normal runs use the smooth shell.
        area_m2 = float(np.count_nonzero(active) * cell_area_m2)
        terrain_volume = float(np.sum((cell_top_z[active] - bottom_z) * cell_area_m2)) if np.any(active) else 0.0
      flat_volume = area_m2 * target_depth
      if name == "Idealised Mooring Corridors":
        paths = [route[:, :2] for _turbine, route, _color in route_paths]
        path_length = float(sum(
            np.sum(np.linalg.norm(np.diff(path, axis=0), axis=1))
            for path in paths if path.shape[0] >= 2
        ))
      else:
        paths, path_length = best_lawnmower(active)
      theoretical = area_m2 / max(50.0, 1.0)
      overhead = max(0.0, 100.0 * (path_length / theoretical - 1.0)) if theoretical > 0.0 else 0.0
      metrics[name] = {
          "area_m2": area_m2,
          "flat_volume_m3": flat_volume,
          "terrain_volume_m3": terrain_volume,
          "relief_adjustment_pct": (
              100.0 * (terrain_volume - flat_volume) / flat_volume
              if flat_volume > 0.0 else float("nan")
          ),
          "path_length_m": path_length,
          "segments": len(paths),
          "overhead_pct": overhead,
          "efficiency_pct": 100.0 / (1.0 + overhead / 100.0),
      }
      scenario_paths[name] = paths

    plotter = pv.Plotter(title="UHRS Seismic Volume & Mooring 3D Viewer", window_size=(1440, 900))
    plotter.set_background("white")
    try:
      plotter.enable_depth_peeling(number_of_peels=8, occlusion_ratio=0.0)
    except Exception:
      pass

    plane_size = max(xmax - xmin, ymax - ymin) + max(50.0, target * 10.0)
    center_x, center_y = (xmin + xmax) * 0.5, (ymin + ymax) * 0.5
    sea_actor = plotter.add_mesh(
        pv.Plane(center=(center_x, center_y, sea_level), direction=(0, 0, 1),
                 i_size=plane_size, j_size=plane_size),
        color="#3a8484", opacity=0.35, label="Sea Level",
    )
    procedural_actor = plotter.add_mesh(
        pv.Plane(center=(center_x, center_y, procedural_seabed), direction=(0, 0, 1),
                 i_size=plane_size, j_size=plane_size),
        color="#8b4513", opacity=0.55, label="Procedural Seabed Fallback",
    )
    procedural_actor.SetVisibility(False)

    terrain_actor = plotter.add_mesh(
        surface_mesh(farm_mask), scalars="Bathymetry", cmap=NAVIA_CMAP,
        opacity=0.90, show_scalar_bar=True,
        scalar_bar_args={
            "title": os.path.basename(raster_path), "color": "black",
            "vertical": False, "n_labels": 3, "position_x": 0.72,
            "position_y": 0.04, "width": 0.24, "height": 0.07,
        },
        label="User Raster Bathymetry",
    )

    # Every comparison scenario is rendered from, and measured by, its own
    # smooth closed terrain shell.  This replaces the coarse cell-edge display
    # that caused stepped full-farm, corridor, triangle, and square volumes.
    volume_actors = {}
    for name, mask, color, opacity in scenario_specs:
      smooth_mesh = smooth_scenarios.get(name, {}).get("mesh")
      if smooth_mesh is not None and smooth_mesh.n_cells:
        volume_actor = plotter.add_mesh(
            smooth_mesh, color=color, opacity=opacity, show_edges=False,
            # Flat cap/wall normals avoid artificial spikes at the join in
            # translucent solids while retaining the smooth XY footprint.
            smooth_shading=False, label=name,
        )
        # Each smooth cap shares its physical elevation with bathymetry. Keep
        # it visible without moving it or changing its reported volume.
        try:
          mapper = volume_actor.GetMapper()
          mapper.SetResolveCoincidentTopologyToPolygonOffset()
          mapper.SetRelativeCoincidentTopologyPolygonOffsetParameters(-1.0, -1.0)
        except Exception:
          pass
        volume_actors[name] = volume_actor
        continue
      mesh = make_volume_mesh(mask)
      if mesh is not None and mesh.n_cells:
        volume_actors[name] = plotter.add_mesh(
            mesh.extract_surface(algorithm="dataset_surface"), color=color, opacity=opacity,
            show_edges=False, label=name,
        )
      else:
        volume_actors[name] = None

    def add_paths(paths, color, label, radius=1.3):
      actors = []
      for index, path in enumerate(paths):
        path = np.asarray(path, dtype=float)
        if path.ndim != 2 or path.shape[0] < 2:
          continue
        points = np.column_stack((path[:, :2], np.full(path.shape[0], sea_level)))
        actors.append(plotter.add_mesh(
            pv.lines_from_points(points).tube(radius=radius, n_sides=8), color=color,
            label=label if index == 0 else "",
        ))
      return actors

    path_actors = {
        "Full Farm Survey Path": add_paths(scenario_paths["Full Farm Area"], "#7c3aed", "Full Farm Survey Vessel Path"),
        "Idealised Mooring Survey Path": add_paths(scenario_paths["Idealised Mooring Corridors"], "#f59e0b", "Idealised Mooring Survey Vessel Path", radius=1.8),
        "Intermediate Triangle Survey Path": add_paths(scenario_paths["Intermediate triangular corridor"], "#2563eb", "Intermediate Triangular Survey Path"),
        "Intermediate Square Survey Path": add_paths(scenario_paths["Intermediate square corridor"], "#16a34a", "Intermediate Square Survey Path"),
    }

    turbine_actors = []
    turbine_points = np.asarray([
        (point[0], point[1], sea_level) for point in turbines.values()
    ], dtype=float)
    if turbine_points.size:
      turbine_actors.append(plotter.add_mesh(
          pv.PolyData(turbine_points).glyph(
              geom=pv.Sphere(radius=8.0, theta_resolution=12, phi_resolution=12),
              scale=False, orient=False,
          ), color="#dc2626", show_edges=True, label="Turbines",
      ))

    mooring_actors = []
    for index, group in enumerate(line_groups.values()):
      points = decimate_polyline(group["points"], maximum_segments=80)
      if points.shape[0] >= 2:
        mooring_actors.append(plotter.add_mesh(
            pv.lines_from_points(points).tube(radius=2.5, n_sides=8),
            color=group["color"], label="Mooring Lines" if index == 0 else "",
        ))

    anchor_actors = []
    for index, group in enumerate(anchor_body_groups.values()):
      points = np.asarray(group["points"], dtype=float)
      record = group["record"]
      centre = np.nanmean(points, axis=0)
      anchor_type = str(record.get("Sub_Type", "")).lower()
      width = as_float(record.get("Anchor_Width_m"), 5.0)
      height = as_float(record.get("Anchor_Height_m"), np.nan)
      if not np.isfinite(height) or height <= 0.0:
        height = as_float(record.get("Cylindrical_Length_m"), base_penetration)
      cylindrical_diameter = as_float(record.get("Cylindrical_Diameter_m"), np.nan)
      if np.isfinite(cylindrical_diameter) and cylindrical_diameter > 0.0:
        mesh = pv.Cylinder(
            center=(centre[0], centre[1], centre[2] - height * 0.5), direction=(0, 0, 1),
            radius=cylindrical_diameter * 0.5, height=height, resolution=24,
        )
      elif any(token in anchor_type for token in ("dea", "gravity", "drag", "stevshark")):
        mesh = pv.Box(bounds=(
            float(np.min(points[:, 0])), float(np.max(points[:, 0])),
            float(np.min(points[:, 1])), float(np.max(points[:, 1])),
            centre[2] - height, centre[2],
        ))
      else:
        radius = max(width * 0.5, 0.5)
        mesh = pv.Cylinder(
            center=(centre[0], centre[1], centre[2] - height * 0.5), direction=(0, 0, 1),
            radius=radius, height=height, resolution=24,
        )
      anchor_actors.append(plotter.add_mesh(
          mesh, color="dimgray", opacity=0.90, show_edges=True,
          label="Anchor Bodies" if index == 0 else "",
      ))

    def add_point_glyph(points, color, label, radius):
      if not points:
        return None
      point_data = np.asarray(points, dtype=float)
      return plotter.add_mesh(
          pv.PolyData(point_data).glyph(
              geom=pv.Sphere(radius=radius, theta_resolution=12, phi_resolution=12),
              scale=False, orient=False,
          ), color=color, label=label,
      )

    point_actors = {
        "Padeyes": add_point_glyph(point_groups["Padeye"], "red", "Anchor Padeyes", 5.0),
        "TDP / DDP": [
            add_point_glyph(point_groups["TDP"], "yellow", "TDP", 5.0),
            add_point_glyph(point_groups["DDP"], "gold", "DDP", 5.0),
        ],
        "Fairleads": add_point_glyph(point_groups["Fairlead"], "#6b7280", "Fairleads", 4.0),
        "Joints": add_point_glyph(point_groups["Joints"], "magenta", "Mooring Joints", 6.0),
    }

    perimeter_z = np.array([
        as_float(record.get("Z_Coord"), procedural_seabed)
        for record in perimeter_records
    ], dtype=float)
    if perimeter_xy.shape[0] >= 2:
      perimeter_points = np.column_stack((
          perimeter_xy, np.where(np.isfinite(perimeter_z), perimeter_z, procedural_seabed)
      ))
      perimeter_points = np.vstack((perimeter_points, perimeter_points[0]))
      perimeter_actor = plotter.add_mesh(
          pv.lines_from_points(perimeter_points).tube(radius=1.0, n_sides=8),
          color="purple", opacity=0.45, label="Farm Perimeter",
      )
    else:
      perimeter_actor = None

    full_metrics = metrics["Full Farm Area"]
    info_text = (
        "--- UHRS SEISMIC SURVEY ---\n"
        f"• Raster: {os.path.basename(raster_path)}\n"
        f"• Turbines: {len(turbines)}\n"
        "• Corridor width: 50.0 m\n"
        "• Planned survey overlap: 0.0%\n"
        "• Survey spacing: 50.0 m\n"
        f"• UHRS target depth: {target_depth:.1f} m "
        f"(2 × max anchor penetration {base_penetration:.1f} m)\n"
        f"• Volume bottom: {bottom_z:.1f} m "
        "(minimum selected-raster seabed − target depth)\n"
        f"• Full farm terrain volume: {full_metrics['terrain_volume_m3']:,.0f} m³\n\n"
        "Use 'Open Seismic Metrics' for scenario and path comparison."
    )
    plotter.add_text(info_text, position="upper_left", font_size=8, color="black", shadow=False)

    def show_metrics_window():
      launch_seismic_metrics_detached(
          metrics, [spec[0] for spec in scenario_specs], target_depth,
      )

    def toggle(actors):
      def callback(enabled):
        for actor in actors if isinstance(actors, list) else [actors]:
          if actor is not None and hasattr(actor, "SetVisibility"):
            actor.SetVisibility(bool(enabled))
      return callback

    controls = [
        ("Sea Level", sea_actor),
        ("Procedural Seabed Fallback", procedural_actor),
        ("User Raster Bathymetry", terrain_actor),
        ("Full Farm Area", volume_actors["Full Farm Area"]),
        ("Idealised Mooring Corridors", volume_actors["Idealised Mooring Corridors"]),
        ("Intermediate triangular corridor", volume_actors["Intermediate triangular corridor"]),
        ("Intermediate square corridor", volume_actors["Intermediate square corridor"]),
        ("Full Farm Survey Path", path_actors["Full Farm Survey Path"]),
        ("Idealised Mooring Survey Path", path_actors["Idealised Mooring Survey Path"]),
        ("Intermediate Triangle Survey Path", path_actors["Intermediate Triangle Survey Path"]),
        ("Intermediate Square Survey Path", path_actors["Intermediate Square Survey Path"]),
        ("Turbines", turbine_actors),
        ("Mooring Lines", mooring_actors),
        ("Anchor Bodies", anchor_actors),
        ("Padeyes", point_actors["Padeyes"]),
        ("TDP / DDP", point_actors["TDP / DDP"]),
        ("Fairleads", point_actors["Fairleads"]),
        ("Joints", point_actors["Joints"]),
        ("Perimeter Fence", perimeter_actor),
    ]
    comparison_volume_labels = {
        "Full Farm Area",
        "Intermediate triangular corridor",
        "Intermediate square corridor",
        "Full Farm Survey Path",
        "Intermediate Triangle Survey Path",
        "Intermediate Square Survey Path",
    }
    y_offset, spacing = 16, 24
    for label, actors in reversed(controls):
      has_actor = any(actor is not None for actor in actors) if isinstance(actors, list) else actors is not None
      if not has_actor:
        continue
      # Opening every translucent comparison solid together obscures the
      # optimised corridor, especially for a dense triad.  Keep the useful
      # gold corridor visible by default; comparison volumes remain optional
      # through the same controls.
      default_visible = (
          label != "Procedural Seabed Fallback"
          and label not in comparison_volume_labels
      )
      toggle(actors)(default_visible)
      plotter.add_checkbox_button_widget(
          toggle(actors), value=default_visible, position=(15, y_offset), size=18,
          border_size=1, color_on="#2563eb", color_off="#94a3b8", background_color="#e2e8f0",
      )
      plotter.add_text(label, position=(40, y_offset + 3), font_size=6, color="black")
      y_offset += spacing

    def set_raster_opacity(value):
      if terrain_actor is not None:
        terrain_actor.GetProperty().SetOpacity(float(value) / 100.0)

    plotter.add_slider_widget(
        set_raster_opacity, rng=(0.0, 100.0), value=90.0, title="Raster opacity (%)",
        pointa=(0.72, 0.15), pointb=(0.96, 0.15), style="modern", fmt="%0.0f",
    )
    metric_button = {"widget": None}

    def metrics_callback(enabled):
      if enabled:
        show_metrics_window()
        try:
          metric_button["widget"].GetRepresentation().SetState(0)
        except Exception:
          pass

    metric_button["widget"] = plotter.add_checkbox_button_widget(
        metrics_callback, value=False, position=(15, min(y_offset + 18, 840)), size=22,
        border_size=2, color_on="#2563eb", color_off="#64748b", background_color="#e2e8f0",
    )
    plotter.add_text(
        "Open Seismic Metrics", position=(46, min(y_offset + 22, 844)),
        font_size=8, color="black",
    )

    for name, item in metrics.items():
      print(
          f"[INFO] {name}: flat-reference={item['flat_volume_m3']:,.0f} m³; "
          f"terrain-aware={item['terrain_volume_m3']:,.0f} m³; "
          f"relief adjustment={item['relief_adjustment_pct']:+.2f}%"
      )
    plotter.add_axes()
    plotter.show_grid()
    if parent_figure is not None:
      parent_figure._seismic_volume_plotter = plotter
      parent_figure._seismic_volume_actors = volume_actors
    plotter.show()
    return True
  except Exception as exc:
    print(f"[ERROR] Seismic Volume Calculation could not be opened: {exc}")
    return False


def launch_seismic_volume_calculation_detached(data, raster_path, export_rows):
  """Open the VTK viewer outside the dashboard's Tk event-loop process.

  PyVista owns a native VTK UI loop.  Running that loop directly from a
  Matplotlib/Tk button callback can corrupt the Python thread state on recent
  Python releases, so the viewer is deliberately started by this same script
  in a separate process.  The short-lived payload contains only the active
  raster and already-generated geometry; no QGIS project is involved.
  """
  script_path = os.path.abspath(__file__)
  if not os.path.isfile(script_path):
    raise RuntimeError("Could not locate Main.py to start the seismic viewer.")

  viewer_data = {
      "anchor_depth": float(data.get("anchor_depth", 7.5)),
      # The viewer needs this layout flag to avoid applying normal-system
      # junction infill between intentionally separate triad lines.
      "triad": bool(data.get("triad", False)),
  }
  payload = {
      "data": viewer_data,
      "raster_path": os.path.abspath(raster_path),
      "export_rows": export_rows,
  }
  payload_file = tempfile.NamedTemporaryFile(
      mode="w", encoding="utf-8", suffix="_seismic_viewer.json",
      prefix="mooring_", delete=False,
  )
  try:
    json.dump(payload, payload_file, default=_seismic_json_default)
    payload_file.close()
  except Exception:
    payload_file.close()
    try:
      os.remove(payload_file.name)
    except OSError:
      pass
    raise

  popen_options = {"cwd": os.path.dirname(script_path)}
  if os.name == "nt":
    # Keep the dashboard tidy while allowing VTK's own visible window.
    popen_options["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
  try:
    subprocess.Popen(
        [sys.executable, script_path, "--seismic-volume-payload", payload_file.name],
        **popen_options,
    )
  except Exception:
    try:
      os.remove(payload_file.name)
    except OSError:
      pass
    raise


def _seismic_json_default(value):
  """Serialise NumPy scalar values retained in in-memory export rows."""
  if isinstance(value, np.generic):
    return value.item()
  if isinstance(value, np.ndarray):
    return value.tolist()
  raise TypeError(f"{type(value).__name__} is not JSON serialisable")


def run_seismic_volume_payload(payload_path):
  """Run the detached PyVista viewer and remove its one-use payload file."""
  try:
    with open(payload_path, "r", encoding="utf-8") as payload_file:
      payload = json.load(payload_file)
    raster_path = payload.get("raster_path")
    export_rows = payload.get("export_rows")
    if not raster_path or not os.path.isfile(raster_path):
      raise FileNotFoundError("The raster selected for Seismic Volume Calculation is unavailable.")
    if not export_rows:
      raise ValueError("The Seismic Volume Calculation payload contains no mooring geometry.")
    return launch_seismic_volume_calculation(
        payload.get("data") or {}, raster_path, export_rows=export_rows,
    )
  except Exception as exc:
    print(f"[ERROR] Seismic Volume Calculation could not be opened: {exc}")
    return False
  finally:
    try:
      os.remove(payload_path)
    except OSError:
      pass


def launch_seismic_metrics_detached(metrics, scenario_names, target_depth):
  """Open metrics outside VTK's callback/event-loop thread."""
  script_path = os.path.abspath(__file__)
  payload_file = tempfile.NamedTemporaryFile(
      mode="w", encoding="utf-8", suffix="_seismic_metrics.json",
      prefix="mooring_", delete=False,
  )
  try:
    json.dump({
        "metrics": metrics,
        "scenario_names": scenario_names,
        "target_depth": float(target_depth),
    }, payload_file, default=_seismic_json_default)
    payload_file.close()
  except Exception:
    payload_file.close()
    try:
      os.remove(payload_file.name)
    except OSError:
      pass
    raise

  popen_options = {"cwd": os.path.dirname(script_path)}
  if os.name == "nt":
    popen_options["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
  try:
    subprocess.Popen(
        [sys.executable, script_path, "--seismic-metrics-payload", payload_file.name],
        **popen_options,
    )
  except Exception:
    try:
      os.remove(payload_file.name)
    except OSError:
      pass
    raise


def show_seismic_metrics_window(metrics, scenario_names, target_depth):
  """Render the UHRS comparison table in the metrics-only process."""
  full_metrics = metrics["Full Farm Area"]
  metric_figure = plt.figure(figsize=(15, 6), facecolor="white")
  try:
    metric_figure.canvas.manager.set_window_title("UHRS Seismic Metrics & Survey Path Comparison")
  except Exception:
    pass
  metric_axis = metric_figure.add_axes([0.02, 0.25, 0.96, 0.68])
  metric_axis.axis("off")
  table_rows = []
  for name in scenario_names:
    item = metrics[name]
    reduction = (
        100.0 * (1.0 - item["terrain_volume_m3"] / full_metrics["terrain_volume_m3"])
        if full_metrics["terrain_volume_m3"] > 0.0 else 0.0
    )
    table_rows.append([
        name,
        f"{item['area_m2'] / 1e6:,.3f}", f"{target_depth:,.1f}",
        f"{item['flat_volume_m3']:,.0f}", f"{item['terrain_volume_m3']:,.0f}",
        f"{item['relief_adjustment_pct']:+,.1f}%", f"{reduction:,.1f}%",
        f"{item['path_length_m'] / 1000.0:,.2f}", str(item["segments"]),
        f"{item['overhead_pct']:,.1f}%",
    ])
  table = metric_axis.table(
      cellText=table_rows,
      colLabels=(
          "Scenario", "Plan Area\n(km²)", "UHRS Target\nDepth (m)",
          "Flat-Reference\nVolume (m³)", "Terrain Survey\nVolume (m³)",
          "Relief\nAdjustment", "Volume\nReduction", "Survey Path\n(km)",
          "Segments", "Route Overlap\nOverhead",
      ),
      loc="center", cellLoc="center",
  )
  table.auto_set_font_size(False)
  table.set_fontsize(8)
  table.scale(1.0, 1.65)
  metric_figure.text(
      0.02, 0.955, "UHRS Seismic Volume & Survey Path Comparison",
      fontsize=13, fontweight="bold", va="top",
  )
  metric_figure.text(
      0.02, 0.07,
      "Terrain-aware survey volume integrates the selected raster down to one fixed bottom datum. "
      "Flat-reference volume is plan area × target depth.  The relief adjustment is the difference between them.\n"
      "The selected raster and the current in-memory mooring export are the only inputs; no QGIS layer is read.",
      fontsize=9, va="bottom",
  )
  plt.show()


def run_seismic_metrics_payload(payload_path):
  """Read a one-use metrics payload, display it, then remove the payload."""
  try:
    with open(payload_path, "r", encoding="utf-8") as payload_file:
      payload = json.load(payload_file)
    show_seismic_metrics_window(
        payload["metrics"], payload["scenario_names"], payload["target_depth"],
    )
    return True
  except Exception as exc:
    print(f"[ERROR] Seismic metrics could not be opened: {exc}")
    return False
  finally:
    try:
      os.remove(payload_path)
    except OSError:
      pass


def normalise_raster_depth_threshold(depth_threshold):
  """Convert the Excel H30 percentage into a fractional +/- tolerance.

  Excel stores a formatted 10% cell as ``0.10``.  A user may also enter ``10``
  directly, so both forms intentionally mean +/-10%.  Values outside 0-100%
  are rejected rather than silently creating an unusable location band.
  """
  try:
    threshold = float(depth_threshold)
  except (TypeError, ValueError):
    return None
  if not np.isfinite(threshold) or threshold < 0.0:
    return None
  if threshold <= 1.0:
    return float(threshold)
  if threshold <= 100.0:
    return float(threshold / 100.0)
  return None


def get_raster_depth_band(design_water_depth, depth_threshold=RASTER_DEPTH_TOLERANCE_FRACTION):
  """Return the inclusive H30 raster-depth suitability band for an Excel target.

  Raster elevations are negative below sea level, whereas this helper works in
  positive water depths.  Keeping the conversion here prevents the picker and
  the generated-farm validation from drifting apart.
  """
  try:
    design_depth = abs(float(design_water_depth))
  except (TypeError, ValueError):
    return None
  tolerance = normalise_raster_depth_threshold(depth_threshold)
  if (
      not np.isfinite(design_depth)
      or design_depth <= 0.0
      or tolerance is None
      or not np.isfinite(tolerance)
      or tolerance > 1.0
  ):
    return None
  return {
      "design_depth": float(design_depth),
      "tolerance_fraction": float(tolerance),
      "threshold_source": "Excel H30",
      "min_depth": float(design_depth * (1.0 - tolerance)),
      "max_depth": float(design_depth * (1.0 + tolerance)),
  }


def assess_raster_depth_band(
    design_water_depth,
    observed_depth,
    depth_threshold=RASTER_DEPTH_TOLERANCE_FRACTION,
    label="Raster position",
    x=None,
    y=None,
):
  """Evaluate one turbine, line or physical anchor against the Excel band.

  This is deliberately a siting gate, not a new mooring solver.  A valid local
  geometry remains available to the renderer even if its raster position is
  unsuitable for the design water depth.
  """
  band = get_raster_depth_band(design_water_depth, depth_threshold)
  try:
    depth = float(observed_depth)
  except (TypeError, ValueError):
    depth = np.nan
  valid_depth = bool(np.isfinite(depth) and depth > 0.0)

  result = {
      "label": str(label),
      "observed_depth": float(depth) if np.isfinite(depth) else np.nan,
      "valid_depth": valid_depth,
      "within_tolerance": False,
      "design_depth": np.nan,
      "min_depth": np.nan,
      "max_depth": np.nan,
  }
  try:
    x_value = float(x)
    if np.isfinite(x_value):
      result["x"] = x_value
  except (TypeError, ValueError):
    pass
  try:
    y_value = float(y)
    if np.isfinite(y_value):
      result["y"] = y_value
  except (TypeError, ValueError):
    pass

  if band is None:
    result["reason"] = "Excel design water depth is unavailable or invalid."
    return result

  result.update(band)
  if not valid_depth:
    result["reason"] = "No valid raster water depth is available."
    return result

  in_band = (
      band["min_depth"] - RASTER_DEPTH_BAND_EPSILON_M
      <= depth
      <= band["max_depth"] + RASTER_DEPTH_BAND_EPSILON_M
  )
  result["within_tolerance"] = bool(in_band)
  if not in_band:
    result["reason"] = (
        f"{depth:.2f} m is outside {band['min_depth']:.2f}-{band['max_depth']:.2f} m "
        f"(Excel target {band['design_depth']:.2f} m +/- {band['tolerance_fraction'] * 100.0:.0f}%)."
    )
  return result


def format_raster_depth_band_failure(check):
  """Return a concise, user-facing failure reason for one raster-depth check."""
  label = str(check.get("label", "Raster position"))
  if check.get("within_tolerance"):
    return ""
  if not check.get("valid_depth"):
    return f"{label}: no valid raster water depth."

  depth = check.get("observed_depth", np.nan)
  lower = check.get("min_depth", np.nan)
  upper = check.get("max_depth", np.nan)
  design = check.get("design_depth", np.nan)
  tolerance = check.get("tolerance_fraction", np.nan)
  try:
    depth, lower, upper, design, tolerance = (
        float(v) for v in (depth, lower, upper, design, tolerance)
    )
  except (TypeError, ValueError):
    depth = lower = upper = design = tolerance = np.nan
  if all(np.isfinite(v) for v in (depth, lower, upper, design, tolerance)):
    return (
        f"{label}: raster depth {depth:.2f} m is outside the permitted "
        f"{lower:.2f}-{upper:.2f} m range "
        f"(Excel design depth {design:.2f} m; H30 +/-{tolerance * 100.0:.2f}%)."
    )
  return f"{label}: raster depth-band check failed."


def raster_line_effective_pass(local_result):
  """Return the solver result combined with the independent site-depth gate."""
  return bool(local_result.get("success")) and bool(
      local_result.get("_raster_depth_band_pass", True)
  )


def get_taut_anchor_seabed_geometry(
    raster_path,
    anchor_x,
    anchor_y,
    fallback_z=None,
    anchor_height=0.0,
    raster_sampler=None,
):
  """Return the authoritative Taut anchor vertical geometry.

  For Taut systems the padeye/head is always exactly on the local raster
  seabed at the physical anchor XY location.  The anchor body then extends
  vertically downward by its physical length.  ``fallback_z`` is used only
  when no raster is available (e.g. base/offline solver runs).
  """
  raster_z = np.nan
  if raster_path:
    try:
      raster_z = float(
          raster_sampler.sample(anchor_x, anchor_y)
          if raster_sampler is not None
          else sample_raster_seabed(raster_path, anchor_x, anchor_y)
      )
    except Exception:
      raster_z = np.nan
  if not np.isfinite(raster_z):
    if fallback_z is None or not np.isfinite(float(fallback_z)):
      raise ValueError(
          f"No valid seabed elevation for Taut anchor at ({float(anchor_x):.3f}, {float(anchor_y):.3f})."
      )
    raster_z = float(fallback_z)
  padeye_z = float(raster_z)
  anchor_top_z = padeye_z
  anchor_bottom_z = anchor_top_z - max(float(anchor_height or 0.0), 0.0)
  return {
      "seabed_z": padeye_z,
      "padeye_z": padeye_z,
      "anchor_top_z": anchor_top_z,
      "anchor_bottom_z": anchor_bottom_z,
  }


def evaluate_dashboard_turbine(data, depth, anchor_db, primary_anc, seabed_z=None, raster_path=None):
  """Re-run the selected mooring model using the raster-derived local water depth.

  All user-selected mooring parameters remain unchanged; only the local seabed
  depth is changed for this turbine/line location.
  """
  if not np.isfinite(depth) or depth <= 0:
    return {"success": False, "status": "FAIL", "msg": "No valid raster seabed depth."}
  try:
    anchor_radius = float(data["anchor_radius"])
    anchor_width = float(data.get("anchor_width", 0.0))
    r_padeye = anchor_radius - (anchor_width / 2.0 if data.get("system_type") == "Taut" else anchor_width)
    common = dict(
        anchor_radius=anchor_radius,
        fairlead_radius=float(r_padeye - data["xf"]),
        fairlead_draft=float(data["fairlead_draft"]),
        num_turbines=1,
        farm_area=data.get("farm_area"),
        num_lines=1,
        center_lat=float(data["center_lat"]),
        center_lon=float(data["center_lon"]),
        buffer_zone=float(data.get("buffer_zone", 50.0)),
        anchor_db=anchor_db,
        primary_anc=primary_anc,
    )
    system = data["system_type"]
    padeye_params = data.get("padeye_params")
    authoritative_seabed_z = (
        float(seabed_z) if seabed_z is not None and np.isfinite(float(seabed_z))
        else -float(depth)
    )
    if system == "Catenary":
      return evaluate_catenary_anchor(
          depth, common["anchor_radius"], float(data["line_length"]),
          float(data["chain_d"]), common["fairlead_radius"], common["fairlead_draft"],
          1, common["farm_area"], 1, common["center_lat"], common["center_lon"],
          common["buffer_zone"], anchor_db, primary_anc, padeye_params=padeye_params,
      )
    if system == "Semi-Taut":
      return evaluate_semi_taut_anchor(
          depth, common["anchor_radius"], float(data["line_length"]),
          float(data["rope_d"]), float(data["chain_d"]), common["fairlead_radius"],
          common["fairlead_draft"], 1, common["farm_area"], 1,
          common["center_lat"], common["center_lon"], common["buffer_zone"],
          anchor_db, primary_anc, padeye_params=padeye_params,
          taut_percentage=float(data.get("taut_percentage", 30.0)),
          upper_equals_lower=bool(data.get("upper_equals_lower", True)),
          lower_joint_pos=float(data.get("lower_joint_pos", 0.50)),
      )
    taut_result = evaluate_taut_anchor(
        depth, common["anchor_radius"], float(data["rope_d"]),
        common["fairlead_radius"], common["fairlead_draft"], 1,
        common["farm_area"], 1, common["center_lat"], common["center_lon"],
        common["buffer_zone"], anchor_db, primary_anc,
        seabed_z=authoritative_seabed_z,
    )
    taut_result["seabed_z"] = authoritative_seabed_z
    taut_result["padeye_z"] = authoritative_seabed_z
    taut_result["anchor_top_z"] = authoritative_seabed_z
    taut_result["anchor_bottom_z"] = authoritative_seabed_z - float(taut_result.get("anchor_height", data.get("anchor_height", 0.0)))
    return taut_result
  except Exception as e:
    return {"success": False, "status": "FAIL", "msg": str(e)}


def terrain_following_line_geometry(
    data, local_result, raster_path, turbine_xy, heading_deg, raster_sampler=None
):
  """Return raster-resting DDP->TDP groundline without altering suspended sections.

  For Semi-Taut lines the raster-following portion is explicitly limited to the
  bottom-chain/ground section.  The smooth transition is also capped before the
  bottom-chain/rope joint so the synthetic rope is never pulled down onto the seabed
  by the display-only terrain replacement.
  """
  if not local_result.get("success"):
    return None

  # The same local result is used by the 2-D profile and the 3-D farm view.
  # Retain its display-only terrain profile so profile navigation, opacity
  # changes, and bridle rendering do not re-sample the raster unnecessarily.
  try:
    _terrain_cache_key = (
        os.path.normcase(os.path.abspath(raster_path)) if raster_path else "",
        round(float(turbine_xy[0]), 6),
        round(float(turbine_xy[1]), 6),
        round(float(heading_deg), 6),
        round(float(data.get("mooring_seabed_clearance", 0.02)), 6),
    )
    _terrain_cache = local_result.get("_terrain_geometry_cache")
    if (
        isinstance(_terrain_cache, dict)
        and _terrain_cache.get("key") == _terrain_cache_key
        and _terrain_cache.get("result") is not None
    ):
      return _terrain_cache["result"]
  except (TypeError, ValueError, OSError):
    _terrain_cache_key = None

  system = local_result.get("system_type", data.get("system_type", "Catenary"))
  if system == "Taut":
    return None

  x_plot = np.asarray(local_result.get("x_plot", []), dtype=float)
  z_plot = np.asarray(local_result.get("z_plot", []), dtype=float)
  sub_x = np.asarray(local_result.get("sub_x", []), dtype=float)
  sub_z = np.asarray(local_result.get("sub_z", []), dtype=float)
  X_td = local_result.get("X_td")
  rope_nodes = local_result.get("rope_nodes")
  if x_plot.size < 2 or z_plot.size != x_plot.size or X_td is None:
    return None
  if sub_x.size < 1 or sub_z.size != sub_x.size:
    return None

  anchor_radius = float(local_result.get("anchor_radius", data.get("anchor_radius", 0.0)))
  anchor_width = float(local_result.get("anchor_width", data.get("anchor_width", 0.0)))
  r_padeye = anchor_radius - anchor_width
  x_ddp = float(sub_x[-1])
  x_td = float(X_td)

  full_x = np.concatenate((sub_x, x_plot[1:]))
  full_z = np.concatenate((sub_z, z_plot[1:]))

  ang = np.radians(float(heading_deg))
  ux, uy = np.cos(ang), np.sin(ang)
  cx, cy = turbine_xy
  radial = r_padeye - full_x
  px = cx + radial * ux
  py = cy + radial * uy
  terrain_z = np.asarray(
      raster_sampler.sample(px, py)
      if raster_sampler is not None
      else sample_raster_seabed(raster_path, px, py),
      dtype=float,
  )

  z_render = full_z.copy()
  clearance = max(0.0, float(data.get("mooring_seabed_clearance", 0.02)))

  # Only the DDP->TDP portion may be replaced by raster ground contact.
  ground = (
      np.isfinite(terrain_z)
      & (full_x >= x_ddp - 1e-9)
      & (full_x <= x_td + 1e-9)
  )
  z_render[ground] = terrain_z[ground] + clearance

  # Smooth transition from the groundline into the suspended lower-chain curve.
  span = max(abs(x_td - x_ddp), 1.0)
  blend = min(max(5.0, 0.10 * span), 20.0)
  if system == "Semi-Taut" and rope_nodes is not None:
    x1_n = float(rope_nodes[0])
    # Never let terrain blending enter the synthetic-rope section.
    blend = min(blend, max(0.0, x1_n - x_td))
  transition = (
      np.isfinite(terrain_z)
      & (full_x >= x_td)
      & (full_x <= x_td + max(blend, 1e-9))
  )
  if np.any(transition) and blend > 0.0:
    t_local = np.clip((full_x[transition] - x_td) / blend, 0.0, 1.0)
    smooth = t_local * t_local * (3.0 - 2.0 * t_local)
    ground_values = terrain_z[transition] + clearance
    z_render[transition] = (
        (1.0 - smooth) * ground_values + smooth * full_z[transition]
    )

  result = {
      "x": full_x,
      "z": z_render,
      "terrain_z": terrain_z,
      "x_ddp": x_ddp,
      "x_td": x_td,
      "r_padeye": r_padeye,
  }
  if rope_nodes is not None:
    result["x_rope_start"] = float(rope_nodes[0])
    result["x_rope_end"] = float(rope_nodes[2])
  if _terrain_cache_key is not None:
    local_result["_terrain_geometry_cache"] = {
        "key": _terrain_cache_key,
        "result": result,
    }
  return result


def _section_curve(x, z, x0, x1):
  """Return the solver curve between x0 and x1, including exact endpoints."""
  x = np.asarray(x, dtype=float)
  z = np.asarray(z, dtype=float)
  if x.size == 0 or z.size != x.size:
    return np.array([]), np.array([])
  lo, hi = sorted((float(x0), float(x1)))
  mask = np.isfinite(x) & np.isfinite(z) & (x >= lo - 1e-9) & (x <= hi + 1e-9)
  xs = x[mask]
  zs = z[mask]
  if xs.size:
    order = np.argsort(xs)
    xs, zs = xs[order], zs[order]
  else:
    xs, zs = np.array([]), np.array([])

  def interp(xq):
    finite = np.isfinite(x) & np.isfinite(z)
    xx, zz = x[finite], z[finite]
    if xx.size < 2:
      return float(np.nan)
    order = np.argsort(xx)
    return float(np.interp(float(xq), xx[order], zz[order]))

  z0 = interp(lo)
  z1 = interp(hi)
  xs = np.concatenate(([lo], xs, [hi]))
  zs = np.concatenate(([z0], zs, [z1]))
  keep = np.concatenate(([True], np.diff(xs) > 1e-10))
  return xs[keep], zs[keep]


def plot_semi_taut_2d_sections(
    ax, data, selected, selected_ground=None, labels=True, line_failed=False
):
  """Plot the original solved Semi-Taut multi-segment profile.

  IMPORTANT:
    The lower-chain/grounded section is ALWAYS retained.  Inverted-bridle
    rendering happens at Joint 1; it must never hide the existing bottom-chain
    or raster-ground section.
  """
  x_plot = np.asarray(selected.get("x_plot", []), dtype=float)
  z_plot = np.asarray(selected.get("z_plot", []), dtype=float)
  sub_x = np.asarray(selected.get("sub_x", []), dtype=float)
  sub_z = np.asarray(selected.get("sub_z", []), dtype=float)
  # A site-depth failure is a siting status, not a geometry failure.  Keep the
  # original solved sections visible, but make the affected profile unambiguous.
  chain_colour = UI_FAIL if line_failed else "black"
  rope_colour = UI_FAIL if line_failed else "deepskyblue"
  rope_outline_colour = "#7a1010" if line_failed else "black"
  rope_nodes = selected.get("rope_nodes")
  if x_plot.size < 2 or z_plot.size != x_plot.size or rope_nodes is None:
    return False

  x1_n, z1_n, x2_n, z2_n = map(float, rope_nodes)
  xf = float(selected.get("xf", x_plot[-1]))
  fairlead_draft = float(
      selected.get("fairlead_draft", data.get("fairlead_draft", 0.0))
  )

  # -------------------------------------------------------------------------
  # BOTTOM CHAIN
  #
  # Use the original solved profile for the complete lower-chain path, but
  # force its final point to the EXACT solver Lower Joint (x1_n, z1_n).
  # This is important for H19=FALSE because the joint location is controlled
  # by H21 and can sit between the sampled solver points.
  # -------------------------------------------------------------------------
  if sub_x.size and sub_z.size and sub_z.size == sub_x.size:
    sx = np.asarray(sub_x, dtype=float)
    sz = np.asarray(sub_z, dtype=float)
    cx, cz = _section_curve(x_plot, z_plot, x_plot[0], x1_n)
    if cx.size:
      # Replace the interpolated endpoint with the exact solver joint.
      cx[0] = float(x_plot[0])
      cz[0] = float(z_plot[0])
      cx[-1] = x1_n
      cz[-1] = z1_n
      bx = np.concatenate((sx, cx[1:]))
      bz = np.concatenate((sz, cz[1:]))
    else:
      bx = np.asarray([sx[-1], x1_n], dtype=float)
      bz = np.asarray([sz[-1], z1_n], dtype=float)
  else:
    bx, bz = _section_curve(x_plot, z_plot, x_plot[0], x1_n)
    if bx.size >= 2:
      bx[0], bz[0] = float(x_plot[0]), float(z_plot[0])
      bx[-1], bz[-1] = x1_n, z1_n

  # Remove duplicate/non-finite points before plotting.
  if bx.size and bz.size == bx.size:
    finite = np.isfinite(bx) & np.isfinite(bz)
    bx, bz = bx[finite], bz[finite]
    if bx.size >= 2:
      order = np.argsort(bx)
      bx, bz = bx[order], bz[order]
      keep = np.r_[True, np.diff(bx) > 1e-10]
      bx, bz = bx[keep], bz[keep]
      # Re-impose the exact joint after duplicate removal.
      bx[-1], bz[-1] = x1_n, z1_n

  # Apply raster ground only over the actual DDP -> TDP interval.
  # The suspended lower chain remains the solved curve all the way to J1.
  if (
      selected_ground is not None
      and bx.size >= 2
      and hasattr(selected_ground, "get")
  ):
    gx = np.asarray(selected_ground.get("x", []), dtype=float)
    gz = np.asarray(selected_ground.get("z", []), dtype=float)
    x_ddp = selected_ground.get("x_ddp")
    x_td = selected_ground.get("x_td")

    if (
        gx.size >= 2
        and gz.size == gx.size
        and x_ddp is not None
        and x_td is not None
    ):
      valid_g = np.isfinite(gx) & np.isfinite(gz)
      gx, gz = gx[valid_g], gz[valid_g]

      if gx.size >= 2:
        order = np.argsort(gx)
        gx, gz = gx[order], gz[order]

        lo = min(float(x_ddp), float(x_td))
        hi = max(float(x_ddp), float(x_td))
        ground_mask = (bx >= lo - 1e-9) & (bx <= hi + 1e-9)

        if np.any(ground_mask):
          bx_ground = bx[ground_mask]
          # np.interp gives the raster-following seabed elevation at the
          # existing solver x positions; no mooring geometry is changed.
          bz_ground = np.interp(
              bx_ground,
              gx,
              gz,
              left=gz[0],
              right=gz[-1],
          )
          bz[ground_mask] = bz_ground

  if bx.size:
    ax.plot(
        bx,
        bz,
        color=chain_colour,
        linewidth=2.6,
        zorder=2,
        label="Bottom Chain" if labels else "",
    )

  # -------------------------------------------------------------------------
  # SYNTHETIC ROPE — exact Lower Joint -> Upper Joint connection
  # -------------------------------------------------------------------------
  rx, rz = _section_curve(x_plot, z_plot, x1_n, x2_n)
  if rx.size >= 2:
    rx[0], rz[0] = x1_n, z1_n
    rx[-1], rz[-1] = x2_n, z2_n
  elif np.isfinite(x1_n) and np.isfinite(z1_n) and np.isfinite(x2_n) and np.isfinite(z2_n):
    rx = np.asarray([x1_n, x2_n], dtype=float)
    rz = np.asarray([z1_n, z2_n], dtype=float)
  if rx.size >= 2:
    ax.plot(
        rx, rz, color=rope_outline_colour, linewidth=6.0, zorder=2,
        label="Synthetic Rope" if labels else "",
    )
    ax.plot(
        rx, rz, color=rope_colour, linewidth=4.0, zorder=3,
        label=(f"Synthetic Rope ({data.get('rope_material', 'Polyester')})" if labels else ""),
    )

  # -------------------------------------------------------------------------
  # TOP CHAIN — exact Upper Joint -> Fairlead connection
  # -------------------------------------------------------------------------
  tx, tz = _section_curve(x_plot, z_plot, x2_n, xf)
  if tx.size >= 2:
    tx[0], tz[0] = x2_n, z2_n
    tx[-1], tz[-1] = xf, fairlead_draft
  elif np.isfinite(x2_n) and np.isfinite(z2_n):
    tx = np.asarray([x2_n, xf], dtype=float)
    tz = np.asarray([z2_n, fairlead_draft], dtype=float)
  if tx.size >= 2:
    ax.plot(
        tx, tz, color=chain_colour, linewidth=2.6, zorder=2,
        label="Top Chain" if labels else "",
    )

  if labels:
    ax.plot(
        x1_n,
        z1_n,
        "D",
        color="magenta",
        markeredgecolor="black",
        markersize=6,
        zorder=6,
        label="Lower Joint",
    )
    ax.plot(
        x2_n,
        z2_n,
        "s",
        color="magenta",
        markeredgecolor="black",
        markersize=6,
        zorder=6,
        label="Upper Joint",
    )
  else:
    ax.plot([x1_n, x2_n], [z1_n, z2_n], linestyle="none")

  return True


def _raster_evaluation_cache_key(data, raster_path, anchor_db, primary_anc):
  """Return the inputs that determine a raster-aware dashboard evaluation."""
  # Keep the cache deliberately narrow: it contains only data that influences
  # the local solver, turbine positions, or H30 depth checks.  Rendering-only
  # fields are excluded so changing a view option does not force hundreds of
  # solver calls.
  input_keys = (
      "system_type", "water_depth", "raster_depth_threshold",
      "num_turbines", "num_lines", "farm_area", "buffer_zone",
      "center_lat", "center_lon", "anchor_radius", "anchor_width",
      "anchor_height", "anchor_depth", "fairlead_draft", "xf",
      "line_length", "chain_d", "rope_d", "taut_percentage",
      "triad", "inter_arm_angle", "angle_between_mooring_arms",
      "upper_equals_lower", "lower_joint_pos", "inverted_bridle",
      "inter_anchor_angle", "padeye_params",
  )
  try:
    normalised_raster_path = (
        os.path.normcase(os.path.abspath(raster_path)) if raster_path else ""
    )
  except (TypeError, OSError):
    normalised_raster_path = str(raster_path or "")
  try:
    raster_stat = os.stat(normalised_raster_path) if normalised_raster_path else None
    raster_fingerprint = (
        int(raster_stat.st_mtime_ns),
        int(raster_stat.st_size),
    ) if raster_stat is not None else None
  except OSError:
    raster_fingerprint = None
  return (
      normalised_raster_path,
      raster_fingerprint,
      str(primary_anc),
      repr(anchor_db.get(primary_anc, {})),
      tuple((key, repr(data.get(key))) for key in input_keys),
  )


def evaluate_raster_configuration(data, raster_path, anchor_db, primary_anc):
  """Run the existing mooring solver for every turbine/line at raster depth.

  This helper deliberately performs no plotting.  It is used when a new
  mooring system is selected so that every anchor type is tested against the
  raster before the user can select it.  The Excel design depth with the H30
  threshold is also used as an inclusive site-depth gate for every generated
  turbine and anchor/padeye location.
  """
  cache_key = _raster_evaluation_cache_key(
      data, raster_path, anchor_db, primary_anc
  )
  cached = data.get("_raster_evaluation_cache")
  if (
      isinstance(cached, dict)
      and cached.get("key") == cache_key
      and isinstance(cached.get("line_results"), list)
  ):
    return cached["line_results"]

  raster_sampler = _RasterSampler(raster_path)
  system_type = data.get("system_type", "Catenary")
  num_turbines = int(data.get("num_turbines", 1))
  num_lines = int(data.get("num_lines", 3))
  farm_area = data.get("farm_area")
  buffer_zone = float(data.get("buffer_zone", 50.0))
  center_lat = float(data["center_lat"])
  center_lon = float(data["center_lon"])
  anchor_radius = float(data["anchor_radius"])
  anchor_width = float(data["anchor_width"])
  origin_x, origin_y = latlon_to_epsg3857(center_lat, center_lon)
  min_spacing = (2.0 * anchor_radius) + buffer_zone
  base_spacing = (np.sqrt((farm_area * 1e6) / num_turbines) if farm_area and farm_area > 0 else min_spacing)
  spacing = max(min_spacing, base_spacing)
  rows, cols = get_optimal_grid_dimensions(num_turbines)
  total_width_x = cols * spacing
  total_height_y = rows * spacing * 0.866
  turbine_coords = []
  for i in range(num_turbines):
    row_idx, col_idx = i // cols, i % cols
    cx = origin_x - (total_width_x / 2.0) + (col_idx + 0.5 * (row_idx % 2)) * spacing
    cy = origin_y - (total_height_y / 2.0) + row_idx * spacing * 0.866
    turbine_coords.append((cx, cy))

  # ``data['water_depth']`` remains the original Excel design depth.  Local
  # results produced below contain raster-derived depths, so never use those
  # values as the tolerance target or every check would trivially pass.
  design_depth = data.get("water_depth", np.nan)
  depth_threshold = data.get(
      "raster_depth_threshold", RASTER_DEPTH_TOLERANCE_FRACTION
  )
  data["_raster_depth_threshold"] = depth_threshold
  data["_raster_depth_band"] = get_raster_depth_band(
      design_depth, depth_threshold
  )

  # The picker validates only the selected farm centre.  Validate every actual
  # generated turbine centre as well, because a farm placed on the edge of the
  # permitted region can otherwise extend into unsuitable water depth.
  turbine_center_z = np.full(len(turbine_coords), np.nan, dtype=float)
  if turbine_coords:
    try:
      _center_z = np.asarray(
          raster_sampler.sample(
              np.asarray([p[0] for p in turbine_coords], dtype=float),
              np.asarray([p[1] for p in turbine_coords], dtype=float),
          ),
          dtype=float,
      ).reshape(-1)
      if _center_z.size == len(turbine_coords):
        turbine_center_z = _center_z
    except Exception:
      pass
  turbine_center_checks = []
  for _ti, ((_cx, _cy), _center_z) in enumerate(zip(turbine_coords, turbine_center_z), start=1):
    _center_depth = -float(_center_z) if np.isfinite(_center_z) else np.nan
    turbine_center_checks.append(assess_raster_depth_band(
        design_depth,
        _center_depth,
        depth_threshold=depth_threshold,
        label=f"T{_ti} turbine centre",
        x=_cx,
        y=_cy,
    ))

  headings = generate_mooring_headings(
      num_lines,
      triad=bool(data.get("triad", False)),
      inter_arm_angle_deg=float(data.get("inter_arm_angle", 0.0)),
      angle_between_mooring_arms_deg=float(
          data.get("angle_between_mooring_arms", 360.0 / max(num_lines, 1))
      ),
  )
  line_results = []
  r_padeye = anchor_radius - (anchor_width / 2.0 if system_type == "Taut" else anchor_width)

  # Read all ordinary padeye elevations in a single raster request.  The
  # previous per-line call opened the raster once for every turbine/line,
  # which became the dominant dashboard load time for large farms.
  padeye_specs = []
  for _i, (_cx, _cy) in enumerate(turbine_coords):
    for _j, _heading in enumerate(headings):
      _angle_rad = np.radians(float(_heading))
      padeye_specs.append((
          _i,
          _j,
          _cx + r_padeye * np.cos(_angle_rad),
          _cy + r_padeye * np.sin(_angle_rad),
      ))
  padeye_raster_z = {}
  if padeye_specs:
    try:
      _padeye_zs = np.asarray(
          raster_sampler.sample(
              np.asarray([spec[2] for spec in padeye_specs], dtype=float),
              np.asarray([spec[3] for spec in padeye_specs], dtype=float),
          ),
          dtype=float,
      ).reshape(-1)
    except Exception:
      _padeye_zs = np.full(len(padeye_specs), np.nan, dtype=float)
    for _spec_index, (_ti, _li, _px, _py) in enumerate(padeye_specs):
      padeye_raster_z[(_ti, _li)] = (
          float(_padeye_zs[_spec_index])
          if _spec_index < _padeye_zs.size and np.isfinite(_padeye_zs[_spec_index])
          else np.nan
      )

  for i, (cx, cy) in enumerate(turbine_coords):
    turbine_lines = []
    for j, heading in enumerate(headings):
      ang = np.radians(float(heading))
      padeye_x = cx + r_padeye * np.cos(ang)
      padeye_y = cy + r_padeye * np.sin(ang)
      raster_z = padeye_raster_z.get((i, j), np.nan)
      local_depth = -float(raster_z) if np.isfinite(raster_z) else np.nan
      local = evaluate_dashboard_turbine(
          data, local_depth, anchor_db, primary_anc,
          seabed_z=(float(raster_z) if np.isfinite(raster_z) else None),
          raster_path=raster_path,
      )
      if system_type == "Taut" and np.isfinite(raster_z):
        local["seabed_z"] = float(raster_z)
        local["padeye_z"] = float(raster_z)
        local["anchor_top_z"] = float(raster_z)
        local["anchor_bottom_z"] = float(raster_z) - float(local.get("anchor_height", data.get("anchor_height", 0.0)))

        # Geometry must remain available even when the engineering feasibility
        # test fails. A failed Taut line is still a valid geometric result for
        # rendering/diagnostics; PASS/FAIL remains authoritative separately.
        if (
            "x_plot" not in local
            or "z_plot" not in local
            or len(np.asarray(local.get("x_plot", []))) < 2
        ):
          try:
            _xf_local = float(local.get("xf", 0.0))
            _z0_local = float(local.get("z_anchor_custom", raster_z))
            _ha_local = float(local.get("HA", np.nan))
            _va_local = float(local.get("VA", np.nan))
            _wf_local = float(
                abs(np.pi * (float(data["rope_d"]) / 2.0) ** 2 *
                    (970.0 - 1025.0) * 9.81)
            )

            if (
                np.isfinite(_xf_local)
                and _xf_local > 0.0
                and np.isfinite(_ha_local)
                and abs(_ha_local) > 1e-9
                and np.isfinite(_va_local)
                and np.isfinite(_wf_local)
                and abs(_wf_local) > 1e-12
            ):
              _a_local = abs(_ha_local) / abs(_wf_local)
              _xv_local = -(
                  abs(_ha_local) / abs(_wf_local)
              ) * np.arcsinh(
                  abs(_va_local) / max(abs(_ha_local), 1e-9)
              )
              _gx = np.linspace(0.0, _xf_local, 500)
              _gz = _z0_local + _a_local * (
                  np.cosh((_gx - _xv_local) / _a_local)
                  - np.cosh(-_xv_local / _a_local)
              )
              _gz[-1] = float(data["fairlead_draft"])
              local["x_plot"] = _gx
              local["z_plot"] = _gz
              local["geometry_available"] = True
          except Exception as _geom_err:
            local["geometry_error"] = str(_geom_err)

      local.update({
          "turbine_index": i + 1, "line_index": j + 1,
          "heading_deg": float(heading),
          "raster_z": float(raster_z) if np.isfinite(raster_z) else np.nan,
          "local_depth": local_depth,
          "anchor_x": float(padeye_x), "anchor_y": float(padeye_y),
      })

      # Preserve the engineering solver result separately from the depth-band
      # siting rule.  The renderer deliberately keeps a locally solved geometry
      # visible even when the selected site is unsuitable.
      local["_raster_solver_pass"] = bool(local.get("success"))
      _depth_checks = [
          dict(turbine_center_checks[i]),
          assess_raster_depth_band(
              design_depth,
              local_depth,
              depth_threshold=depth_threshold,
              label=f"T{i+1}L{j+1} line anchor / solver reference",
              x=padeye_x,
              y=padeye_y,
          ),
      ]

      # Refresh only the inverted-bridle load wrapper at this raster depth.
      # The user's original Catenary/Semi-Taut/Taut solver is not modified.
      if bool(data.get("inverted_bridle", False)) and local.get("success"):
        try:
          local["inverted_bridle"] = True
          local["triad"] = bool(data.get("triad", False))
          local["upper_equals_lower"] = bool(data.get("upper_equals_lower", True))
          local["lower_joint_pos"] = float(data.get("lower_joint_pos", 0.5))
          local["inter_anchor_angle"] = float(data.get("inter_anchor_angle", 30.0))
          _bg = calculate_inverted_bridle_geometry(
              local,
              float(heading),
              (float(cx), float(cy)),
              raster_path=raster_path,
              raster_sampler=raster_sampler,
          )
          local["bridle_geometry_template"] = _bg
          local["bridle_loads"] = calculate_inverted_bridle_loads(local, _bg)
          local["bridle_anchor_results"] = {}
          for _bi, _br in enumerate((local.get("bridle_loads") or {}).get("branches", []), start=1):
            local["bridle_anchor_results"][f"Branch {_bi}"] = evaluate_anchor_capacities(
              anchor_db, [primary_anc], _br["HA"], _br["VA"]
            ).get(primary_anc, {})
        except Exception as _ble:
          local["bridle_geometry_error"] = str(_ble)

      # A bridle has two physical anchor/padeye positions, neither of which is
      # the virtual centreline reference used by the one-line solver.  Test both
      # real locations as well so their raster depth can fail the owning line
      # and turbine even when the virtual centreline point is acceptable.
      _bridle_anchors = []
      _bridle_template = local.get("bridle_geometry_template") or {}
      for _branch_index, _branch_anchor in enumerate(_bridle_template.get("anchors", []), start=1):
        try:
          _branch_x = float(_branch_anchor["x"])
          _branch_y = float(_branch_anchor["y"])
          if np.isfinite(_branch_x) and np.isfinite(_branch_y):
            _bridle_anchors.append((_branch_index, _branch_anchor, _branch_x, _branch_y))
        except (KeyError, TypeError, ValueError):
          continue
      if _bridle_anchors:
        try:
          _branch_zs = np.asarray(
              raster_sampler.sample(
                  np.asarray([item[2] for item in _bridle_anchors], dtype=float),
                  np.asarray([item[3] for item in _bridle_anchors], dtype=float),
              ),
              dtype=float,
          ).reshape(-1)
        except Exception:
          _branch_zs = np.full(len(_bridle_anchors), np.nan, dtype=float)
        for _offset, (_branch_index, _branch_anchor, _branch_x, _branch_y) in enumerate(_bridle_anchors):
          _branch_z = _branch_zs[_offset] if _offset < _branch_zs.size else np.nan
          _branch_depth = -float(_branch_z) if np.isfinite(_branch_z) else np.nan
          _suffix = chr(ord("a") + _branch_index - 1)
          _branch_anchor["raster_z"] = float(_branch_z) if np.isfinite(_branch_z) else np.nan
          _branch_anchor["local_depth"] = _branch_depth
          _branch_check = assess_raster_depth_band(
              design_depth,
              _branch_depth,
              depth_threshold=depth_threshold,
              label=f"T{i+1}L{j+1}{_suffix} physical bridle anchor",
              x=_branch_x,
              y=_branch_y,
          )
          _branch_check["branch_index"] = _branch_index
          _depth_checks.append(_branch_check)

      _depth_failures = [
          check for check in _depth_checks if not bool(check.get("within_tolerance"))
      ]
      local["_raster_turbine_depth_check"] = dict(turbine_center_checks[i])
      local["_raster_depth_band_checks"] = _depth_checks
      local["_raster_depth_band_pass"] = not _depth_failures
      local["_raster_depth_band_failures"] = _depth_failures
      local["_raster_depth_band_reason"] = "; ".join(
          reason for reason in (format_raster_depth_band_failure(check) for check in _depth_failures)
          if reason
      )
      local["_raster_effective_pass"] = raster_line_effective_pass(local)
      turbine_lines.append(local)
    line_results.append(turbine_lines)
  data["_raster_path"] = raster_path
  data["_raster_anchor_elevations"] = {
      f"{int(local.get('turbine_index', 0))}:{int(local.get('line_index', 0))}": float(local.get("seabed_z", local.get("raster_z", np.nan)))
      for lines in line_results for local in lines
      if np.isfinite(float(local.get("seabed_z", local.get("raster_z", np.nan))))
  }
  data["_raster_turbine_depth_checks"] = turbine_center_checks
  failure_reasons = []
  for i, center_check in enumerate(turbine_center_checks):
    if not bool(center_check.get("within_tolerance")):
      reason = format_raster_depth_band_failure(center_check)
      if reason:
        failure_reasons.append(f"T{i+1}: {reason}")
  for i, lines in enumerate(line_results):
    for j, local in enumerate(lines):
      if not raster_line_effective_pass(local):
        reasons = []
        if not bool(local.get("_raster_solver_pass", local.get("success"))):
          reasons.append(local.get("msg", "Local raster-depth feasibility check failed."))
        for check in local.get("_raster_depth_band_failures", []):
          # The turbine-centre error was added once above; do not repeat it for
          # every one of the turbine's mooring lines.
          if check.get("label") == f"T{i+1} turbine centre":
            continue
          reason = format_raster_depth_band_failure(check)
          if reason:
            reasons.append(reason)
        if reasons:
          failure_reasons.append(
              f"T{i+1}L{j+1}: " + " | ".join(dict.fromkeys(str(r) for r in reasons))
          )
  data["_raster_failure_reasons"] = failure_reasons
  data["_raster_turbine_results"] = [
      {
          "turbine_index": i + 1,
          "success": bool(lines) and all(raster_line_effective_pass(v) for v in lines),
          "solver_success": bool(lines) and all(bool(v.get("success")) for v in lines),
          "depth_band_pass": bool(lines) and all(bool(v.get("_raster_depth_band_pass", True)) for v in lines),
          "turbine_depth_check": dict(turbine_center_checks[i]) if i < len(turbine_center_checks) else {},
          "lines": lines,
      }
      for i, lines in enumerate(line_results)
  ]
  # Keep the engineering solver result separate from the Excel H30 siting
  # result.  A depth-only failure must remain drawable so the user can inspect
  # the affected line in the unified 2-D/3-D profile.
  data["_raster_solver_system_pass"] = bool(line_results) and all(
      bool(lines) and all(bool(v.get("success")) for v in lines)
      for lines in line_results
  )
  data["_raster_depth_band_failure"] = any(
      not bool(v.get("_raster_depth_band_pass", True))
      for lines in line_results for v in lines
  )
  data["_raster_system_pass"] = bool(line_results) and all(
      bool(lines) and all(raster_line_effective_pass(v) for v in lines) for lines in line_results
  )
  data["_raster_evaluation_cache"] = {
      "key": cache_key,
      "line_results": line_results,
  }
  raster_sampler.close()
  return line_results


def build_raster_depth_summary(data):
  """Return a compact, turbine-centred summary of the raster depth checks.

  Each physical bridle anchor is still checked independently, but shared
  turbine-centre failures are reported once.  This keeps the Failure Criteria
  report readable when several branches of the same turbine fail together.
  """
  data = data or {}
  depth_band = data.get("_raster_depth_band") or {}
  try:
    band_valid = all(
        np.isfinite(float(depth_band.get(key, np.nan)))
        for key in ("design_depth", "min_depth", "max_depth")
    )
  except (TypeError, ValueError):
    band_valid = False
  if band_valid:
    band_text = (
        f"Excel H30 depth band: {float(depth_band['design_depth']):.1f} m target; "
        f"{float(depth_band['min_depth']):.1f}–{float(depth_band['max_depth']):.1f} m allowed "
        f"(+/-{float(depth_band.get('tolerance_fraction', np.nan)) * 100.0:.2f}%)."
    )
  else:
    band_text = "Excel H30 depth band is unavailable; local raster positions cannot pass suitability."

  turbines = []
  for position, turbine in enumerate(data.get("_raster_turbine_results") or [], start=1):
    turbine_number = int(turbine.get("turbine_index", position) or position)
    lines = turbine.get("lines") or []
    depths = []
    issues = []
    seen_issues = set()
    for local in lines:
      try:
        local_depth = float(local.get("local_depth", np.nan))
        if np.isfinite(local_depth):
          depths.append(local_depth)
      except (TypeError, ValueError):
        pass
      for check in local.get("_raster_depth_band_failures", []) or []:
        label = str(check.get("label", "Raster position"))
        prefix = f"T{turbine_number}"
        if label.startswith(prefix):
          label = label[len(prefix):].strip()
        label = label.replace(" line anchor / solver reference", " anchor")
        label = label.replace(" physical bridle anchor", " anchor")
        try:
          observed_depth = float(check.get("observed_depth", np.nan))
        except (TypeError, ValueError):
          observed_depth = np.nan
        item = (
            f"{label}: {observed_depth:.1f} m"
            if bool(check.get("valid_depth")) and np.isfinite(observed_depth)
            else f"{label}: no raster depth"
        )
        if item not in seen_issues:
          seen_issues.add(item)
          issues.append(item)
    if not bool(turbine.get("solver_success", True)) and not issues:
      issues.append("local mooring solver failure")
    depth_text = (
        f"{min(depths):.1f}–{max(depths):.1f} m"
        if depths else "no line-anchor depth"
    )
    turbines.append({
        "number": turbine_number,
        "passed": bool(turbine.get("success", False)),
        "depth_text": depth_text,
        "issues": issues,
    })
  return {"band_text": band_text, "turbines": turbines}


def _dedupe_mooring_legend(ax):
  """Clean the profile/farm legend into one entry per physical concept.

  Inverted bridles deliberately render multiple graphical objects (trunk,
  selected branch, taut black outline/blue centreline, etc.).  Those are
  implementation layers, not separate legend items.
  """
  handles, labels = ax.get_legend_handles_labels()
  if not handles:
    return

  semi_taut_2d = bool(getattr(ax, "_semi_taut_2d_legend", False))

  def canonical(label):
    text = str(label or "").strip()
    low = text.lower()
    if not text or low == "_nolegend_":
      return None

    # Semi-Taut has three physically distinct line materials/sections in the
    # 2-D profile.  Keep them separate in the key rather than collapsing every
    # chain/rope object into a generic Mooring Line entry.
    if semi_taut_2d:
      if "synthetic rope" in low or "synthetic section" in low or "middle rope" in low:
        return "Middle Rope"
      if "top chain" in low:
        return "Top Chain"
      if "bottom chain" in low:
        return "Bottom Chain"
      if "mooring line" in low or "chain section" in low or "trunk" in low:
        return "Bottom Chain"

    if ("mooring" in low or "trunk" in low or "synthetic rope" in low
        or "synthetic section" in low or "chain section" in low
        or (len(text) >= 5 and text[0].upper() == "T" and "L" in text and
            text[-1:].lower() in ("a", "b"))):
      return "Mooring Line"
    if "padeye" in low:
      return "Padeye"
    if "anchor body" in low or low == "anchor" or low.endswith(" anchor"):
      return "Anchor"
    if "bottom chain / rope joint" in low or low == "bottom joint" or "lower joint" in low:
      return "Lower Joint"
    if "rope / top chain joint" in low or low == "upper joint" or low == "top joint":
      return "Upper Joint"
    if "touchdown point" in low or low == "tdp":
      return "Touchdown Point (TDP)"
    if "down dip point" in low or low == "ddp":
      return "Down Dip Point (DDP)"
    return text

  out_h, out_l = [], []
  seen = set()
  for h, l in zip(handles, labels):
    cl = canonical(l)
    if cl is None or cl in seen:
      continue
    seen.add(cl)
    out_h.append(h)
    out_l.append(cl)
  leg = ax.legend(out_h, out_l, loc="upper right", fontsize=5.5)
  try:
    leg.set_zorder(10000)
  except Exception:
    pass



def _bridle_branch_2d_x(branch):
  """Return the original solver radial coordinate for the 2-D profile.

  The 2-D dashboard is a vertical/radial profile, not a plan-view projection.
  H25 changes the branch azimuth in 3-D, but it must not rescale the horizontal
  profile distance.  Using the branch padeye-to-joint 3-D span here stretches
  the lower branch and makes it inconsistent with the solved trunk, TDP and
  Fairlead coordinates.
  """
  px = np.asarray(branch.get("profile_x", []), dtype=float)
  if px.size >= 2 and np.all(np.isfinite(px)):
    return px
  return np.asarray(branch.get("profile_x_2d", []), dtype=float)


def _normalise_bridle_2d_branch(branch):
  """Return a clean, monotonic 2-D branch profile without changing solver data.

  The bridle geometry is solved in the original radial section.  For very long
  lines, the Lower Joint can fall between sparse solver samples; this helper
  sorts/deduplicates the displayed branch samples so interpolation and endpoint
  snapping are deterministic.
  """
  px = np.asarray(branch.get("profile_x", branch.get("profile_x_2d", [])), dtype=float)
  pz = np.asarray(branch.get("z", branch.get("profile_z", [])), dtype=float)
  if px.size < 2 or pz.size != px.size:
    return np.array([], dtype=float), np.array([], dtype=float)
  good = np.isfinite(px) & np.isfinite(pz)
  px, pz = px[good], pz[good]
  if px.size < 2:
    return np.array([], dtype=float), np.array([], dtype=float)
  order = np.argsort(px, kind="mergesort")
  px, pz = px[order], pz[order]
  keep = np.r_[True, np.diff(px) > 1e-10]
  return px[keep], pz[keep]


def _taut_bridle_2d_display_profile(branch, trunk_x, trunk_z, joint_x, joint_z):
  """Return a C1-continuous display-only profile for a Taut bridle branch.

  Each Taut branch starts at its own raster-derived padeye elevation.  The
  required endpoint remap can slightly change the branch's terminal slope even
  though the shared solver trunk retains its original tangent.  This helper
  replaces only the final display samples with a short cubic-Hermite transition
  that matches that trunk tangent.  It never changes the solved geometry,
  tensions, 3-D branch coordinates, or export values.
  """
  bx, bz = _normalise_bridle_2d_branch(branch)
  if bx.size < 4:
    return bx, bz

  bx = np.asarray(bx, dtype=float).copy()
  bz = np.asarray(bz, dtype=float).copy()
  joint_x = float(joint_x)
  joint_z = float(joint_z)
  bx[-1] = joint_x
  bz[-1] = joint_z

  tx = np.asarray(trunk_x, dtype=float)
  tz = np.asarray(trunk_z, dtype=float)
  good = np.isfinite(tx) & np.isfinite(tz)
  tx, tz = tx[good], tz[good]
  if tx.size < 2:
    return bx, bz
  order = np.argsort(tx, kind="mergesort")
  tx, tz = tx[order], tz[order]
  keep = np.r_[True, np.diff(tx) > 1e-10]
  tx, tz = tx[keep], tz[keep]
  if tx.size < 2:
    return bx, bz

  # Use the original solved trunk's local tangent at the exact lower joint.
  right = int(np.searchsorted(tx, joint_x, side="right"))
  if right <= 0:
    right = 1
  elif right >= tx.size:
    right = tx.size - 1
  dx_trunk = float(tx[right] - tx[right - 1])
  if not np.isfinite(dx_trunk) or abs(dx_trunk) <= 1e-12:
    return bx, bz
  trunk_slope = float((tz[right] - tz[right - 1]) / dx_trunk)
  if not np.isfinite(trunk_slope):
    return bx, bz

  # Keep most of the physical lower branch unchanged; only replace a short
  # end section so it joins the original taut trunk without a rendered corner.
  n_blend = int(max(6, round(0.12 * bx.size)))
  n_blend = min(n_blend, max(3, bx.size // 2))
  start_idx = bx.size - n_blend
  blend_x0 = float(bx[start_idx])
  blend_span = joint_x - blend_x0
  if not np.isfinite(blend_span) or blend_span <= 1e-10:
    return bx, bz
  if start_idx > 0:
    dx_start = float(bx[start_idx] - bx[start_idx - 1])
    branch_slope = (
        float((bz[start_idx] - bz[start_idx - 1]) / dx_start)
        if abs(dx_start) > 1e-12 else trunk_slope
    )
  else:
    branch_slope = float((joint_z - bz[0]) / blend_span)
  if not np.isfinite(branch_slope):
    branch_slope = trunk_slope

  t = np.clip((bx[start_idx:] - blend_x0) / blend_span, 0.0, 1.0)
  h00 = (2.0 * t**3) - (3.0 * t**2) + 1.0
  h10 = t**3 - (2.0 * t**2) + t
  h01 = (-2.0 * t**3) + (3.0 * t**2)
  h11 = t**3 - t**2
  bz[start_idx:] = (
      h00 * bz[start_idx]
      + h10 * blend_span * branch_slope
      + h01 * joint_z
      + h11 * blend_span * trunk_slope
  )
  bz[-1] = joint_z
  return bx, bz


def _draw_dashboard_anchor_3d(
    ax,
    padeye_x,
    padeye_y,
    padeye_z,
    heading_deg,
    system_type,
    primary_anc,
    anchor_width,
    anchor_height,
    anchor_depth,
    seabed_z,
    data,
    show_legend=False,
    opacity=1.0,
):
  """Draw one physical padeye and its matching anchor body in the 3-D dashboard."""
  # The dashboard slider supports true zero opacity, including physical
  # anchor and padeye geometry rendered through this helper.
  opacity = float(np.clip(opacity, 0.0, 1.0))
  values = (padeye_x, padeye_y, padeye_z, heading_deg)
  if not all(np.isfinite(float(value)) for value in values):
    return

  padeye_x = float(padeye_x)
  padeye_y = float(padeye_y)
  padeye_z = float(padeye_z)
  heading_rad = np.radians(float(heading_deg))
  ux, uy = np.cos(heading_rad), np.sin(heading_rad)
  anc_type_upper = str(primary_anc).upper()
  is_taut = system_type == "Taut"
  is_cylinder = any(
      key in anc_type_upper for key in ("DRIVEN", "DRILLED", "SUCTION", "PILE")
  )

  ax.scatter(
      [padeye_x], [padeye_y], [padeye_z],
      color="red", marker="o", s=24,
      edgecolors="black", linewidths=0.6, alpha=opacity, zorder=42,
      label="Padeye" if show_legend else "",
  )

  if is_cylinder:
    cyl_dia = float(anchor_width)
    cyl_len = float(anchor_height)
  else:
    box_w = float(anchor_width)
    box_l = float(anchor_width)
    box_h = float(anchor_height)

  if is_taut:
    cx_anc, cy_anc = padeye_x, padeye_y
    z_top = padeye_z
    z_bot = padeye_z - (cyl_len if is_cylinder else box_h)
  elif is_cylinder and any(
      key in anc_type_upper for key in ("SUCTION", "DRIVEN", "DRILLED")
  ):
    frac = float(
        (data.get("padeye_params") or {}).get(
            primary_anc, {}
        ).get("position_fraction", 0.50)
    )
    local_depth = max(-float(seabed_z), 0.0)
    z_bot, z_top, _ = get_pile_anchor_vertical_geometry(
        local_depth, anchor_depth, anchor_height, padeye_fraction=frac
    )
    cx_anc = padeye_x + (cyl_dia / 2.0) * ux
    cy_anc = padeye_y + (cyl_dia / 2.0) * uy
  elif is_cylinder:
    cx_anc = padeye_x + (cyl_dia / 2.0) * ux
    cy_anc = padeye_y + (cyl_dia / 2.0) * uy
    z_top = padeye_z + (cyl_len / 2.0)
    z_bot = padeye_z - (cyl_len / 2.0)
  else:
    cx_anc = padeye_x + (box_l / 2.0) * ux
    cy_anc = padeye_y + (box_l / 2.0) * uy
    z_top = padeye_z + (box_h / 2.0)
    z_bot = padeye_z - (box_h / 2.0)

  body_label = f"Anchor Body ({primary_anc})" if show_legend else ""
  if is_cylinder:
    theta = np.linspace(0, 2 * np.pi, 24)
    x_circle = cx_anc + (cyl_dia / 2.0) * np.cos(theta)
    y_circle = cy_anc + (cyl_dia / 2.0) * np.sin(theta)
    ax.plot(
        x_circle, y_circle, np.full_like(x_circle, z_bot),
        color="purple", lw=1.2, alpha=0.8 * opacity, zorder=40,
        label=body_label,
    )
    ax.plot(
        x_circle, y_circle, np.full_like(x_circle, z_top),
        color="purple", lw=1.2, alpha=0.8 * opacity, zorder=40,
    )
    for idx_strut in range(0, len(theta), 6):
      ax.plot(
          [x_circle[idx_strut], x_circle[idx_strut]],
          [y_circle[idx_strut], y_circle[idx_strut]],
          [z_bot, z_top], color="purple", lw=1.2, alpha=0.8 * opacity,
          zorder=40,
      )
  else:
    vx, vy = -uy, ux
    dx, dy = box_l / 2.0, box_w / 2.0
    x_box = np.asarray([
        cx_anc + dx * ux + dy * vx, cx_anc - dx * ux + dy * vx,
        cx_anc - dx * ux - dy * vx, cx_anc + dx * ux - dy * vx,
        cx_anc + dx * ux + dy * vx,
    ], dtype=float)
    y_box = np.asarray([
        cy_anc + dx * uy + dy * vy, cy_anc - dx * uy + dy * vy,
        cy_anc - dx * uy - dy * vy, cy_anc + dx * uy - dy * vy,
        cy_anc + dx * uy + dy * vy,
    ], dtype=float)
    ax.plot(
        x_box, y_box, np.full_like(x_box, z_bot),
        color="purple", lw=1.2, alpha=0.8 * opacity, zorder=40,
        label=body_label,
    )
    ax.plot(
        x_box, y_box, np.full_like(x_box, z_top),
        color="purple", lw=1.2, alpha=0.8 * opacity, zorder=40,
    )
    for x_corner, y_corner in zip(x_box[:-1], y_box[:-1]):
      ax.plot(
          [x_corner, x_corner], [y_corner, y_corner], [z_bot, z_top],
          color="purple", lw=1.2, alpha=0.8 * opacity, zorder=40,
      )


def _bridle_render_joint_2d(bridle_geometry, selected_branch_index=1):
  """Return the exact Lower Joint used by the displayed 2-D bridle curves.

  The solver/3-D joint remains authoritative.  The 2-D renderer additionally
  snaps its Z value to the selected displayed branch at the same X coordinate,
  preventing the rare long-line case where the marker is a fraction of a pixel
  away from the rendered curve and appears disconnected.
  """
  if not bridle_geometry or not bridle_geometry.get("enabled"):
    return None
  try:
    jx = float(bridle_geometry["joint_local"][0])
    jz = float(bridle_geometry["joint_local"][1])
  except Exception:
    return None

  profiles = bridle_geometry.get("profiles") or []
  if profiles:
    try:
      idx = max(0, min(int(selected_branch_index or 1) - 1, len(profiles) - 1))
      px, pz = _normalise_bridle_2d_branch(profiles[idx])
      if px.size >= 2:
        jx_render = float(np.clip(jx, px[0], px[-1]))
        jz_render = float(np.interp(jx_render, px, pz))
        return jx_render, jz_render
    except Exception:
      pass
  return jx, jz


def _bridle_trunk_shift_x(branch):
  """Horizontal 2D shift applied to the original suspended trunk."""
  px = np.asarray(branch.get("profile_x", []), dtype=float)
  px2 = _bridle_branch_2d_x(branch)
  if px.size >= 2 and px2.size == px.size:
    return float(px2[-1] - px[-1])
  return 0.0


def _project_semitaur_bridle_2d(branch, selected, bridle_geometry, turbine_xy, heading_deg, r_padeye, fairlead_radius):
  """Compatibility helper: return the standard solver radial profile.

  The 2-D dashboard is intentionally not a plan-view projection. H25 affects
  only the 3-D branch geometry.
  """
  try:
    px = np.asarray(branch.get("profile_x", []), dtype=float)
    pz = np.asarray(branch.get("profile_z", branch.get("z", [])), dtype=float)
    if px.size >= 2 and pz.size == px.size:
      return {"x": px.copy(), "z": pz.copy()}
  except Exception:
    pass
  return None


def render_terrain_dashboard(
    ax1, ax2, data, raster_path, anchor_db, primary_anc, profile_index=0,
    turbine_opacity=1.0,
):
  """Render the dashboard using the ORIGINAL mooring/anchor rendering.

  The only per-line change is the water depth supplied to the existing
  Catenary/Semi-Taut/Taut evaluator.  The original padeye equations, anchor
  geometry, line styles, Semi-Taut segmentation and plotting conventions are
  deliberately retained.
  """
  ax1.clear()
  ax2.clear()
  # Keep the turbine/mooring assets above the raster in 3-D.  Axes3D normally
  # recomputes z-order from depth, which can bury a valid line or padeye inside
  # the raster surface at some view angles.
  try:
    ax2.computed_zorder = False
  except Exception:
    pass
  # A full farm view samples the same raster hundreds of times.  Keep a
  # single read handle open for this render, then close it before returning.
  render_raster_sampler = _RasterSampler(raster_path)
  turbine_opacity = float(np.clip(turbine_opacity, 0.0, 1.0))

  def _asset_alpha(base=1.0):
    return float(np.clip(float(base) * turbine_opacity, 0.0, 1.0))

  system_type = data.get("system_type", "Catenary")
  num_turbines = int(data.get("num_turbines", 1))
  num_lines = int(data.get("num_lines", 3))
  farm_area = data.get("farm_area")
  buffer_zone = float(data.get("buffer_zone", 50.0))
  center_lat = float(data["center_lat"])
  center_lon = float(data["center_lon"])
  anchor_radius = float(data["anchor_radius"])
  anchor_width = float(data["anchor_width"])
  anchor_height = float(data["anchor_height"])
  anchor_depth = float(data["anchor_depth"])
  fairlead_draft = float(data["fairlead_draft"])
  primary_anc = primary_anc
  # Horizontal padeye radius is unchanged. For driven/drilled piles the vertical
  # padeye elevation is handled independently so the anchor body is not moved.
  r_padeye = (
      anchor_radius - (anchor_width / 2.0)
      if system_type == "Taut"
      else anchor_radius - anchor_width
  )

  origin_x, origin_y = latlon_to_epsg3857(center_lat, center_lon)
  min_spacing = (2.0 * anchor_radius) + buffer_zone
  base_spacing = (
      np.sqrt((farm_area * 1e6) / num_turbines)
      if farm_area and farm_area > 0 else min_spacing
  )
  spacing = max(min_spacing, base_spacing)
  rows, cols = get_optimal_grid_dimensions(num_turbines)
  total_width_x = cols * spacing
  total_height_y = rows * spacing * 0.866

  turbine_coords = []
  for i in range(num_turbines):
    row_idx, col_idx = i // cols, i % cols
    cx = origin_x - (total_width_x / 2.0) + (col_idx + 0.5 * (row_idx % 2)) * spacing
    cy = origin_y - (total_height_y / 2.0) + row_idx * spacing * 0.866
    turbine_coords.append((cx, cy))

  headings = generate_mooring_headings(
      num_lines,
      triad=bool(data.get("triad", False)),
      inter_arm_angle_deg=float(data.get("inter_arm_angle", 0.0)),
      angle_between_mooring_arms_deg=float(
          data.get("angle_between_mooring_arms", 360.0 / max(num_lines, 1))
      ),
  )
  line_results = []

  # Per-line raster depth and viability.  The solver is unchanged; only its
  # local water-depth input comes from the raster at the original padeye.
  line_results = evaluate_raster_configuration(data, raster_path, anchor_db, primary_anc)

  # Select the actual turbine/line/bridle branch requested by the profile navigator.
  bridle_navigation = bool(data.get("inverted_bridle", False))
  base_profile_count = max(num_turbines * num_lines, 1)
  total_profiles = max(base_profile_count * (2 if bridle_navigation else 1), 1)
  profile_index = int(profile_index) % total_profiles
  if bridle_navigation:
    base_profile_index = profile_index // 2
    selected_branch_index = (profile_index % 2) + 1
  else:
    base_profile_index = profile_index
    selected_branch_index = None
  selected_turbine_index = base_profile_index // max(num_lines, 1)
  selected_line_index = base_profile_index % max(num_lines, 1)
  selected = line_results[selected_turbine_index][selected_line_index]

  # All plotting below uses the selected result from the ORIGINAL solver.
  # The raster only supplies its local depth; it does not alter rendering.
  water_depth = float(selected.get("local_depth", selected.get("water_depth", np.nan)))
  if system_type == "Taut":
    selected_seabed_z = float(selected.get("seabed_z", selected.get("raster_z", -water_depth)))
    selected["seabed_z"] = selected_seabed_z
    selected["padeye_z"] = selected_seabed_z
    selected["z_anchor_custom"] = selected_seabed_z
    selected["anchor_top_z"] = selected_seabed_z
    selected["anchor_bottom_z"] = selected_seabed_z - float(selected.get("anchor_height", anchor_height))
  # Use the selected line's own solver result for subsurface length.
  # This keeps the T1L1/T1L2/... profile navigation independent and avoids
  # relying on an undefined function-level L_sub variable.
  L_sub = float(selected.get("L_sub", 0.0))
  x_plot = np.asarray(selected.get("x_plot", []), dtype=float)
  z_plot = np.asarray(selected.get("z_plot", []), dtype=float)
  xf = float(selected.get("xf", data.get("xf", 0.0)))
  X_td = selected.get("X_td")
  sub_x = np.asarray(selected.get("sub_x", []), dtype=float)
  sub_z = np.asarray(selected.get("sub_z", []), dtype=float)
  rope_nodes = selected.get("rope_nodes")
  for _bk in (
      "triad",
      "upper_equals_lower",
      "lower_joint_pos",
      "inverted_bridle",
      "inter_anchor_angle",
      "system_type",
      "padeye_z",
      "z_anchor_custom",
  ):
    if _bk not in selected and _bk in data:
      selected[_bk] = data[_bk]
  selected["bridle_geometry"] = None
  fairlead_draft = float(selected.get("fairlead_draft", fairlead_draft))
  anchor_width = float(selected.get("anchor_width", anchor_width))
  anchor_height = float(selected.get("anchor_height", anchor_height))
  anchor_depth = float(selected.get("anchor_depth", anchor_depth))
  if system_type == "Taut":
    anchor_depth = get_taut_effective_penetration_depth(anchor_depth, anchor_height)
  anchor_radius = float(selected.get("anchor_radius", anchor_radius))
  r_padeye = (
      anchor_radius - (anchor_width / 2.0)
      if system_type == "Taut"
      else anchor_radius - anchor_width
  )

  # Raster-resting groundline geometry for the selected Catenary/Semi-Taut line.
  selected_ground = terrain_following_line_geometry(
      data, selected, raster_path,
      turbine_coords[selected_turbine_index],
      headings[selected_line_index],
      raster_sampler=render_raster_sampler,
  ) if system_type != "Taut" else None

  if bool(selected.get("inverted_bridle", data.get("inverted_bridle", False))):
    selected["bridle_geometry"] = prepare_selected_inverted_bridle(
        selected,
        headings[selected_line_index],
        turbine_coords[selected_turbine_index],
        raster_path=raster_path,
        selected_ground=selected_ground,
        raster_sampler=render_raster_sampler,
    )

  # Selected bridle geometry is now fully prepared and can safely be referenced by all rendering paths.
  bridle_2d = selected.get("bridle_geometry") if isinstance(selected, dict) else None
  selected_bridle_anchor = None
  selected_bridle_profile = None
  if (
      bridle_navigation
      and bridle_2d
      and bridle_2d.get("enabled")
      and selected_branch_index in (1, 2)
  ):
    _selected_branch_zero_index = selected_branch_index - 1
    _bridle_anchors = bridle_2d.get("anchors", [])
    _bridle_profiles = bridle_2d.get("profiles", [])
    if len(_bridle_anchors) > _selected_branch_zero_index:
      selected_bridle_anchor = _bridle_anchors[_selected_branch_zero_index]
    if len(_bridle_profiles) > _selected_branch_zero_index:
      selected_bridle_profile = _bridle_profiles[_selected_branch_zero_index]

  # A Taut bridle has two real padeyes, one at each physical branch anchor.
  # Derive their 2-D/profile references once here, before drawing the raster,
  # Z-plane, marker, anchor body and infobox.  This prevents a selected a/b
  # branch from mixing its anchor with the centreline line result.
  selected_taut_bridle_padeye_z = np.nan
  selected_taut_bridle_anchor_bottom_z = np.nan
  if system_type == "Taut" and selected_bridle_anchor is not None:
    try:
      selected_taut_bridle_padeye_z = float(
          selected_bridle_anchor.get(
              "padeye_z", selected_bridle_anchor.get("z", np.nan)
          )
      )
    except (AttributeError, TypeError, ValueError):
      pass
    try:
      selected_taut_bridle_anchor_bottom_z = float(
          selected_bridle_anchor.get("anchor_bottom_z", np.nan)
      )
    except (AttributeError, TypeError, ValueError):
      pass
    if (
        np.isfinite(selected_taut_bridle_padeye_z)
        and not np.isfinite(selected_taut_bridle_anchor_bottom_z)
    ):
      selected_taut_bridle_anchor_bottom_z = (
          selected_taut_bridle_padeye_z - float(anchor_height)
      )

  # -------------------------------------------------------------------------
  # 2D PROFILE — ORIGINAL RENDERING STYLE, with raster seabed underneath.
  # -------------------------------------------------------------------------
  # Do not suppress a valid solved profile merely because its local site is
  # outside Excel H30.  Red is reserved for that siting failure; normal solver
  # failures continue to follow the existing unavailable-profile path.
  profile_depth_failed = not bool(selected.get("_raster_depth_band_pass", True))
  profile_chain_colour = UI_FAIL if profile_depth_failed else "black"
  profile_rope_colour = UI_FAIL if profile_depth_failed else "deepskyblue"
  profile_rope_outline_colour = "#7a1010" if profile_depth_failed else "black"
  ax1.axhline(0, color="blue", linestyle="--", linewidth=1.5,
              zorder=1, label="Sea Level ($z = 0$ m)")

  if selected.get("success") and np.isfinite(water_depth):
    # Sample the actual raster along the original 2D profile path.
    # Use adaptive sampling for long mooring lines so the displayed seabed
    # does not become under-resolved while leaving the solved mooring geometry
    # completely unchanged.  A target spacing of <= 0.5 m is used, capped to
    # keep the renderer responsive for very long lines.
    _render_span = max(float(abs(xf)), 1.0)
    _ground_samples = int(np.clip(np.ceil(_render_span / 0.5) + 1.0, 500, 4000))
    s_ground = np.linspace(0.0, _render_span, _ground_samples)
    ang0 = np.radians(float(headings[selected_line_index]))
    cx0, cy0 = turbine_coords[selected_turbine_index]
    if bridle_navigation and selected_bridle_profile is not None:
      _bp = selected_bridle_profile
      _bx = np.asarray(_bp.get("x", []), dtype=float)
      _by = np.asarray(_bp.get("y", []), dtype=float)
      _bt = np.asarray(_bp.get("terrain_z", []), dtype=float)
      _bpx2 = _bridle_branch_2d_x(_bp)
      _branch_has_xy = (
          _bpx2.size >= 2
          and _bx.size == _bpx2.size
          and _by.size == _bpx2.size
      )

      # Taut deliberately does not use the grounded-line terrain transform,
      # but its two physical bridle anchors still sit on independently sampled
      # raster elevations.  Sample the selected branch's actual XY trace here
      # so the visible seabed meets its selected padeye, rather than falling
      # back to the centreline heading.
      if system_type == "Taut" and _branch_has_xy:
        try:
          _bt = np.asarray(render_raster_sampler.sample(_bx, _by), dtype=float)
          _bp["terrain_z"] = _bt
        except Exception:
          _bt = np.asarray([], dtype=float)

      if _bpx2.size >= 2 and _bt.size == _bpx2.size:
        _branch_ground_span = max(float(abs(_bpx2[-1] - _bpx2[0])), 1.0)
        _branch_ground_samples = int(
            np.clip(np.ceil(_branch_ground_span / 0.5) + 1.0, 500, 4000)
        )
        _branch_s_ground = np.linspace(
            float(_bpx2[0]), float(_bpx2[-1]), _branch_ground_samples
        )
        _branch_seabed = np.interp(_branch_s_ground, _bpx2, _bt)

        if system_type == "Taut" and float(xf) > float(_branch_s_ground[-1]) + 1e-9:
          # Above the lower joint the taut trunk is still the original
          # centreline section.  Join its terrain trace onto the physical
          # branch trace so the seabed remains continuous all the way to the
          # fairlead without fabricating an H25 displacement in the 2-D view.
          _trunk_span = max(float(xf) - float(_branch_s_ground[-1]), 1.0)
          _trunk_samples = int(
              np.clip(np.ceil(_trunk_span / 0.5) + 1.0, 80, 4000)
          )
          _trunk_s_ground = np.linspace(
              float(_branch_s_ground[-1]), float(xf), _trunk_samples
          )
          _trunk_wx = cx0 + (r_padeye - _trunk_s_ground) * np.cos(ang0)
          _trunk_wy = cy0 + (r_padeye - _trunk_s_ground) * np.sin(ang0)
          _trunk_seabed = np.asarray(
              render_raster_sampler.sample(_trunk_wx, _trunk_wy), dtype=float
          )
          s_ground = np.concatenate((_branch_s_ground, _trunk_s_ground[1:]))
          seabed_profile = np.concatenate((_branch_seabed, _trunk_seabed[1:]))
        else:
          s_ground = _branch_s_ground
          seabed_profile = _branch_seabed
      else:
        wx = cx0 + (r_padeye - s_ground) * np.cos(ang0)
        wy = cy0 + (r_padeye - s_ground) * np.sin(ang0)
        seabed_profile = np.asarray(
            render_raster_sampler.sample(wx, wy), dtype=float
        )
    else:
      wx = cx0 + (r_padeye - s_ground) * np.cos(ang0)
      wy = cy0 + (r_padeye - s_ground) * np.sin(ang0)
      seabed_profile = np.asarray(
          render_raster_sampler.sample(wx, wy), dtype=float
      )
    valid_sb = np.isfinite(seabed_profile)
    if np.any(valid_sb):
      ax1.plot(s_ground[valid_sb], seabed_profile[valid_sb],
               color="saddlebrown", linestyle="-", linewidth=3.5,
               zorder=1, label="Seabed (Raster)")
    else:
      ax1.axhline(-water_depth, color="saddlebrown", linestyle="-",
                  linewidth=3.5, zorder=1, label=f"Seabed ($z = {-water_depth:.1f}$ m)")

    # Maximum penetration reference: for every mooring system this is the
    # actual bottom of the anchor below the local seabed.  Taut anchors use the
    # physical anchor length because the padeye/head is at the seabed surface;
    # Catenary/Semi-Taut retain their Excel-defined penetration depth.
    effective_penetration_depth = (
        get_taut_effective_penetration_depth(anchor_depth, anchor_height)
        if system_type == "Taut" else max(float(anchor_depth), 0.0)
    )
    _taut_reference_z = (
        float(selected.get("padeye_z", selected.get("seabed_z", -water_depth)))
        if system_type == "Taut" else -water_depth
    )
    _selected_taut_bridle = (
        system_type == "Taut"
        and np.isfinite(selected_taut_bridle_padeye_z)
    )
    if _selected_taut_bridle:
      _taut_reference_z = float(selected_taut_bridle_padeye_z)
    z_anchor_bottom = (
        float(selected_taut_bridle_anchor_bottom_z)
        if _selected_taut_bridle and np.isfinite(selected_taut_bridle_anchor_bottom_z)
        else _taut_reference_z - effective_penetration_depth
    )
    ax1.axhline(
        z_anchor_bottom, color="purple", linestyle=":", linewidth=1.5,
        zorder=1, label="_nolegend_"
    )
    # Place the penetration annotation near the right side of the actual
    # rendered profile.  For long lines, scaling directly from xf can put the
    # label on top of the line or outside the useful plot area.
    x_annot = float(xf) * 0.96
    if np.isfinite(xf) and xf != 0.0:
      x_annot = float(xf - 0.04 * abs(xf))
    _penetration_label = (
        "Branch Anchor Bottom / Max Penetration"
        if _selected_taut_bridle else "Max Penetration"
    )
    ax1.text(
        x_annot, z_anchor_bottom + 0.6,
        f"{_penetration_label}: z = {z_anchor_bottom:.2f} m",
        color="purple", fontsize=7.0, fontweight="bold",
        ha="right", va="bottom", zorder=60,
        bbox=dict(boxstyle="round,pad=0.18", facecolor="white", alpha=0.75, edgecolor="none"),
    )

    # ORIGINAL line rendering, unchanged in colour/weights/segmentation.
    # A bridle branch may have a different raster-derived padeye elevation
    # from the centreline solver reference.  Keep this explicit coordinate so
    # the 2-D padeye and anchor body can follow the selected branch endpoint.
    anchor_x_2d = 0.0
    taut_anchor_bottom_2d = np.nan
    if system_type == "Taut":
      z_anchor = float(selected.get("padeye_z", selected.get("seabed_z", -water_depth)))
      if np.isfinite(selected_taut_bridle_padeye_z):
        z_anchor = float(selected_taut_bridle_padeye_z)
      taut_anchor_bottom_2d = z_anchor - float(anchor_height)
      if np.isfinite(selected_taut_bridle_anchor_bottom_z):
        taut_anchor_bottom_2d = float(selected_taut_bridle_anchor_bottom_z)

      bridle_2d = selected.get("bridle_geometry")
      if bridle_2d and bridle_2d.get("enabled"):
        _joint_render = _bridle_render_joint_2d(
            bridle_2d, selected_branch_index
        )
        if _joint_render is None:
          _jx_orig = float(bridle_2d["joint_local"][0])
          _jz = float(bridle_2d["joint_local"][1])
        else:
          _jx_orig, _jz = _joint_render

        # IMPORTANT: the taut 2-D profile is a vertical/radial section.
        # The H25 bridle angle is a plan-view (3-D) property and must NOT shift
        # the Fairlead horizontally in the 2-D section.  The previous renderer
        # shifted the entire Fairlead->Lower Joint trunk by the branch projection,
        # which moved the line endpoint away from the actual Fairlead marker.
        _jx = _jx_orig

        # Trunk: Fairlead -> Lower Joint.  Use the original solved taut profile
        # in its original radial x-coordinate so its endpoint remains exactly at
        # the same Fairlead (xf, fairlead_draft) used by the solver and marker.
        _tm = x_plot >= _jx - 1e-8
        _tx = np.concatenate(([_jx], x_plot[_tm]))
        _tz = np.concatenate(([_jz], z_plot[_tm]))
        keep = np.r_[True, np.diff(_tx) > 1e-10]
        _tx, _tz = _tx[keep], _tz[keep]
        # In an inverted-bridle Taut system the lower/trunk section uses the
        # exact same blue styling as the rest of the taut line.  Do not add a
        # separate black lower-section representation.
        ax1.plot(_tx, _tz, color=profile_rope_outline_colour, linewidth=6.0, zorder=2)
        ax1.plot(_tx, _tz, color=profile_rope_colour, linewidth=4.0, zorder=3,
                 label=f"Taut Mooring Trunk ({data.get('rope_material', 'HMPE')})")

        # Two bridle branches.  In a vertical 2-D section their plan separation
        # projects onto the same profile, so both use the same solved elevation
        # curve while the 3-D view carries the actual H25 separation.
        for _bi, _branch in enumerate(bridle_2d.get("profiles", []), start=1):
          # The profile navigator addresses physical branches independently
          # (T1L1a, T1L1b, ...).  Showing both here makes a selected T1L1a
          # profile look like two lower taut lines even though it represents
          # only its one selected branch.
          if bridle_navigation and _bi != selected_branch_index:
            continue
          _bx, _bz = _taut_bridle_2d_display_profile(
              _branch, x_plot, z_plot, _jx, _jz
          )
          if _bx.size >= 2:
            _bx = np.asarray(_bx, dtype=float).copy()
            _bz = np.asarray(_bz, dtype=float).copy()
            _bx[-1] = float(_jx)
            _bz[-1] = float(_jz)
            # The branch's first point is its physical padeye.  Use the exact
            # same point for the marker/body below instead of the centreline
            # anchor location, so TnLna and TnLnb redraw independently.
            if bridle_navigation and _bi == selected_branch_index:
              anchor_x_2d = float(_bx[0])
              z_anchor = (
                  float(selected_taut_bridle_padeye_z)
                  if np.isfinite(selected_taut_bridle_padeye_z)
                  else float(_bz[0])
              )
              if not np.isfinite(selected_taut_bridle_anchor_bottom_z):
                taut_anchor_bottom_2d = z_anchor - float(anchor_height)
            _label = f"T{selected_turbine_index + 1}L{selected_line_index + 1}{'a' if _bi == 1 else 'b'}"
            _branch_label = _label if (not bridle_navigation or _bi == selected_branch_index) else "_nolegend_"
            ax1.plot(_bx, _bz, color=profile_rope_outline_colour, linewidth=6.0, zorder=2)
            ax1.plot(
                _bx, _bz, color=profile_rope_colour, linewidth=4.0, zorder=3,
                label=_branch_label
            )
        ax1.plot(_jx, _jz, "D", color="magenta", markeredgecolor="black",
                 markersize=6, zorder=50, label="Lower Joint")
      else:
        ax1.plot(x_plot, z_plot, color=profile_rope_outline_colour, linewidth=6.0, zorder=2)
        ax1.plot(x_plot, z_plot, color=profile_rope_colour, linewidth=4.0, zorder=3,
                 label=f"Taut Mooring Line ({data.get('rope_material', 'HMPE')})")
      ax1.plot(anchor_x_2d, z_anchor, "ro", markeredgecolor="black", markersize=5,
               zorder=5, label="Padeye")
    else:
      if bridle_2d and bridle_2d.get("enabled") and bridle_navigation:
        # Use exactly the physical padeye belonging to the selected branch.
        _bi = max(1, min(int(selected_branch_index or 1), 2))
        _anchors = bridle_2d.get("anchors", [])
        if len(_anchors) >= _bi and np.isfinite(float(_anchors[_bi - 1].get("z", np.nan))):
          z_anchor = float(_anchors[_bi - 1]["z"])
        else:
          z_anchor = float(selected.get("padeye_z", sub_z[0] if sub_z.size else -water_depth))
      elif any(k in primary_anc.upper() for k in ("SUCTION", "DRIVEN", "DRILLED", "PILE")):
        _, _, z_anchor = get_pile_anchor_vertical_geometry(
            water_depth, anchor_depth, anchor_height,
            padeye_fraction=float((data.get("padeye_params") or {}).get(primary_anc, {}).get("position_fraction", 0.50))
        )
      else:
        z_anchor = sub_z[0] if sub_z is not None and len(sub_z) > 0 else -water_depth
      ax1.plot(0.0, z_anchor, "ro", markeredgecolor="black", markersize=5,
               zorder=50, label="Padeye")

      if (
          system_type == "Semi-Taut"
          and rope_nodes is not None
          and not (bridle_2d and bridle_2d.get("enabled"))
      ):
        plot_semi_taut_2d_sections(
            ax1, data, selected, selected_ground=selected_ground, labels=True,
            line_failed=profile_depth_failed,
        )
      elif rope_nodes is not None and not (bridle_2d and bridle_2d.get("enabled")):
        x1_n, z1_n, x2_n, z2_n = rope_nodes
        mask_bot = x_plot <= x1_n
        bot_x = (np.concatenate((sub_x, x_plot[mask_bot][1:] if np.any(mask_bot) else []))
                 if sub_x is not None else x_plot[mask_bot])
        bot_z = (np.concatenate((sub_z, z_plot[mask_bot][1:] if np.any(mask_bot) else []))
                 if sub_z is not None else z_plot[mask_bot])
        ax1.plot(bot_x, bot_z, color=profile_chain_colour, linewidth=2.5, zorder=2, label="Mooring Line Profile (Chain)")
        ax1.plot([x1_n, x2_n], [z1_n, z2_n], color=profile_rope_outline_colour, linewidth=6.0, zorder=2)
        ax1.plot([x1_n, x2_n], [z1_n, z2_n], color=profile_rope_colour, linewidth=4.0, zorder=3,
                 label=f"Synthetic Section ({data.get('rope_material', 'Polyester')})")
        mask_top = x_plot >= x2_n
        if np.any(mask_top):
          ax1.plot(np.concatenate(([x2_n], x_plot[mask_top], [xf])),
                   np.concatenate(([z2_n], z_plot[mask_top], [fairlead_draft])),
                   color=profile_chain_colour, linewidth=2.5, zorder=2)
        ax1.plot(x1_n, z1_n, "D", color="magenta", markeredgecolor="black", markersize=5, zorder=50, label="Bottom Joint")
        ax1.plot(x2_n, z2_n, "s", color="magenta", markeredgecolor="black", markersize=5, zorder=50, label="Top Joint")
      else:
        # Semi-Taut uses the ORIGINAL continuous solved profile for the 2-D section.
        # The inverted bridle changes the 3-D plan-view arrangement and load path,
        # but both branches have the same vertical 2-D projection.  Reuse the
        # standard Semi-Taut renderer here so the subsurface, grounded DDP->TDP,
        # suspended bottom chain, middle rope and top chain are never stitched
        # together from different coordinate sources.
        if system_type == "Semi-Taut" and rope_nodes is not None:
          plot_semi_taut_2d_sections(
              ax1, data, selected,
              selected_ground=selected_ground,
              labels=True,
              line_failed=profile_depth_failed,
          )
        elif bridle_2d and bridle_2d.get("enabled") and system_type == "Catenary":
          _joint_render = _bridle_render_joint_2d(
              bridle_2d, selected_branch_index
          )
          if _joint_render is None:
            _jx_orig = float(bridle_2d["joint_local"][0])
            _jz = float(bridle_2d["joint_local"][1])
          else:
            _jx_orig, _jz = _joint_render
          _jx = float(_jx_orig)
          _trunk_shift = 0.0

          # Trunk: Lower Joint -> Fairlead, retaining the original Catenary
          # curve and its existing seabed interaction.
          if selected_ground is not None:
            full_x = np.asarray(selected_ground["x"], dtype=float)
            full_z = np.asarray(selected_ground["z"], dtype=float)
          else:
            full_x = np.concatenate((sub_x, x_plot[1:])) if sub_x is not None else x_plot
            full_z = np.concatenate((sub_z, z_plot[1:])) if sub_z is not None else z_plot
          _m = full_x >= _jx_orig - 1e-8
          _full_x = np.concatenate(([_jx], full_x[_m] + _trunk_shift))
          _full_z = np.concatenate(([_jz], full_z[_m]))
          keep = np.r_[True, np.diff(_full_x) > 1e-10]
          _full_x, _full_z = _full_x[keep], _full_z[keep]
          ax1.plot(_full_x, _full_z, color=profile_chain_colour, linewidth=2.5, zorder=2,
                   label="Catenary Trunk")

          # Two lower branches.  In 2-D they share the same vertical projection;
          # their real separation is represented in the 3-D farm view.
          for _bi, _branch in enumerate(bridle_2d.get("profiles", []), start=1):
            if bridle_navigation and _bi != selected_branch_index:
              continue
            _bx, _bz = _normalise_bridle_2d_branch(_branch)
            if _bx.size >= 2:
              # Force the rendered branch to terminate on the exact same 2-D
              # joint used by the trunk/marker.  This only changes plotting
              # endpoints; it never mutates the solved branch geometry.
              _bx = np.asarray(_bx, dtype=float).copy()
              _bz = np.asarray(_bz, dtype=float).copy()
              _bx[-1] = float(_jx)
              _bz[-1] = float(_jz)
              _label = f"T{selected_turbine_index + 1}L{selected_line_index + 1}{'a' if _bi == 1 else 'b'}"
              ax1.plot(
                  _bx, _bz, color=profile_chain_colour, linewidth=2.5, zorder=2,
                  label=_label if (not bridle_navigation or _bi == selected_branch_index) else "_nolegend_"
              )
          ax1.plot(_jx, _jz, "D", color="magenta", markeredgecolor="black",
                   markersize=6, zorder=50, label="Lower Joint")
        else:
          if selected_ground is not None:
            full_x = selected_ground["x"]
            full_z = selected_ground["z"]
          else:
            full_x = np.concatenate((sub_x, x_plot[1:])) if sub_x is not None else x_plot
            full_z = np.concatenate((sub_z, z_plot[1:])) if sub_z is not None else z_plot
          ax1.plot(full_x, full_z, color=profile_chain_colour, linewidth=2.5, zorder=2,
                   label="Mooring Line Profile")
          if bridle_2d and bridle_2d.get("enabled") and system_type == "Semi-Taut":
            for _bi, _branch in enumerate(bridle_2d.get("profiles", []), start=1):
              if bridle_navigation and _bi != selected_branch_index:
                continue
              _bx = np.asarray(_bridle_branch_2d_x(_branch), dtype=float)
              _bz = np.asarray(_branch.get("z", []), dtype=float)
              if _bx.size >= 2 and _bz.size == _bx.size:
                _label = f"T{selected_turbine_index + 1}L{selected_line_index + 1}{'a' if _bi == 1 else 'b'}"
                ax1.plot(
                    _bx, _bz, color=profile_chain_colour, linewidth=2.5, zorder=3,
                    label=_label if (not bridle_navigation or _bi == selected_branch_index) else "_nolegend_"
                )

      # Each system-specific renderer above now draws its bridle trunk and the
      # selected physical branch.  Do not add a second generic bridle overlay:
      # it duplicated the lower section in the 2-D profiler, particularly for
      # a selected Taut branch such as T1L1a.

      # For an inverted bridle the TDP/DDP belong to the selected physical
      # branch, not the midpoint between branches.
      _selected_bridle_profile = selected_bridle_profile

      if _selected_bridle_profile is not None:
        # Use the selected branch explicitly. Do not rely on the loop-local
        # `_branch` variable from an earlier rendering loop; it may not exist
        # when navigation selects a single branch.
        _orig_px = np.asarray(_selected_bridle_profile.get("profile_x", []), dtype=float)
        _branch_z = np.asarray(
            _selected_bridle_profile.get("z", _selected_bridle_profile.get("profile_z", [])),
            dtype=float,
        )
        if _orig_px.size >= 2 and _branch_z.size == _orig_px.size:
          if sub_x is not None and len(sub_x):
            _ddp_orig = float(np.asarray(sub_x, dtype=float)[-1])
            _ddp_z = float(np.interp(_ddp_orig, _orig_px, _branch_z))
            ax1.plot(_ddp_orig, _ddp_z, "s", color="yellow", markeredgecolor="black", markersize=6, zorder=50, label="Down Dip Point (DDP)")
          if X_td is not None:
            _tdp_x = float(X_td)
            _tdp_z = float(np.interp(_tdp_x, _orig_px, _branch_z))
            ax1.plot(_tdp_x, _tdp_z, "o", color="yellow", markeredgecolor="black", markersize=6, zorder=50, label="Touchdown Point (TDP)")
      else:
        if selected_ground is not None:
          ddp_z_render = float(np.interp(
              selected_ground["x_ddp"], selected_ground["x"], selected_ground["z"]
          ))
          ax1.plot(selected_ground["x_ddp"], ddp_z_render, "s", color="yellow",
                   markeredgecolor="black", markersize=6, zorder=6,
                   label="Down Dip Point (DDP)")
        elif sub_x is not None and sub_z is not None and len(sub_x) > 0:
          ax1.plot(sub_x[-1], sub_z[-1], "s", color="yellow", markeredgecolor="black",
                   markersize=6, zorder=6, label="Down Dip Point (DDP)")

    is_top_mid = system_type == "Taut"
    if (not is_top_mid) and any(k in primary_anc.upper() for k in ("SUCTION", "DRIVEN", "DRILLED", "PILE")):
      rect_x = -anchor_width
      rect_z = -water_depth - anchor_depth
    else:
      rect_x = (
          anchor_x_2d - anchor_width / 2.0
          if is_top_mid else -anchor_width
      )
      rect_z = (
          float(taut_anchor_bottom_2d)
          if is_top_mid and np.isfinite(taut_anchor_bottom_2d)
          else z_anchor - anchor_height if is_top_mid
          else z_anchor - (anchor_height / 2.0)
      )
    ax1.add_patch(patches.Rectangle((rect_x, rect_z), anchor_width, anchor_height,
                                    linewidth=1.5, edgecolor="black", facecolor="gray",
                                    zorder=4, label="Anchor"))
    # Keep the 2-D Fairlead on the original solved radial profile.  The
    # bridle azimuth is represented only in the 3-D farm geometry.
    _fairlead_x_render = float(xf)
    ax1.plot(_fairlead_x_render, fairlead_draft, "go", markeredgecolor="black", markersize=5,
             zorder=4, label="Fairlead/Turbine")

    if X_td is not None and not (bridle_navigation and bridle_2d and bridle_2d.get("enabled")):
      # Put the TDP exactly on the displayed mooring line.
      x_td_plot = float(X_td)
      line_x_td = np.asarray(x_plot, dtype=float)
      line_z_td = np.asarray(z_plot, dtype=float)

      if system_type != "Taut" and sub_x is not None and sub_z is not None:
        sx = np.asarray(sub_x, dtype=float)
        sz = np.asarray(sub_z, dtype=float)
        if sx.size and sz.size:
          line_x_td = np.concatenate((sx, line_x_td[1:]))
          line_z_td = np.concatenate((sz, line_z_td[1:]))

      finite = np.isfinite(line_x_td) & np.isfinite(line_z_td)
      if selected_ground is not None:
        xq = float(np.clip(x_td_plot, selected_ground["x"][0], selected_ground["x"][-1]))
        z_td = float(np.interp(xq, selected_ground["x"], selected_ground["z"]))
        x_td_plot = xq
      elif np.count_nonzero(finite) >= 2:
        lx_td = line_x_td[finite]
        lz_td = line_z_td[finite]
        order = np.argsort(lx_td)
        lx_td = lx_td[order]
        lz_td = lz_td[order]
        xq = float(np.clip(x_td_plot, lx_td[0], lx_td[-1]))
        z_td = float(np.interp(xq, lx_td, lz_td))
        x_td_plot = xq
      else:
        z_td = -water_depth

      ax1.plot(x_td_plot, z_td, "o", color="yellow", markeredgecolor="black",
               markersize=6, zorder=50, label="Touchdown Point (TDP)")

  # Original dashboard information box, retained for the selected 2D line.
  selected_info_data = dict(selected)
  _info_profile_label = f"T{selected_turbine_index + 1}L{selected_line_index + 1}"
  if bridle_navigation and selected_branch_index in (1, 2):
    _suffix = "a" if selected_branch_index == 1 else "b"
    _branch_name = f"T{selected_turbine_index + 1}L{selected_line_index + 1}{_suffix}"
    _branch_loads = (selected.get("bridle_loads") or {}).get("branches", [])
    _branch_load = _branch_loads[selected_branch_index - 1] if len(_branch_loads) >= selected_branch_index else {}
    _info_profile_label = _branch_name
    selected_info_data.update({
        "profile_label": _branch_name,
        "T_padeye": _branch_load.get("tension", selected.get("T_padeye", data.get("T_padeye", np.nan))),
        "HA": _branch_load.get("HA", selected.get("HA", data.get("HA", 0.0))),
        "VA": _branch_load.get("VA", selected.get("VA", data.get("VA", 0.0))),
        "line_fos_padeye": _branch_load.get("line_fos", selected.get("line_fos_padeye", data.get("line_fos_padeye", 0.0))),
        "line_fos": _branch_load.get("line_fos", selected.get("line_fos", data.get("line_fos", 0.0))),
    })
  selected_info_data.update({
      "system_type": system_type,
      "primary_anc": primary_anc,
      "anchor_radius": anchor_radius,
      "line_length": data.get("line_length", 0.0),
      "num_lines": num_lines,
      "num_turbines": num_turbines,
      "water_depth": water_depth,
      "fairlead_draft": data.get("fairlead_draft", fairlead_draft),
      "anchor_width": anchor_width,
      "anchor_height": anchor_height,
      "anchor_depth": anchor_depth,
      "L_sub": L_sub,
      "padeye_params": data.get("padeye_params"),
      "chain_d": data.get("chain_d"),
      "rope_d": data.get("rope_d"),
      "anchor_results": selected.get("anchor_results", data.get("anchor_results")),
      "min_geometric_length": selected.get("min_geometric_length", data.get("min_geometric_length", np.nan)),
      "min_geometric_main_length": selected.get("min_geometric_main_length", data.get("min_geometric_main_length", np.nan)),
      "excess_line": selected.get("excess_line", data.get("excess_line", np.nan)),
      "ground_ratio": selected.get("ground_ratio", data.get("ground_ratio", 0.0)),
      "grounded_surface_length": selected.get("grounded_surface_length", data.get("grounded_surface_length", selected.get("L_bot", data.get("L_bot", 0.0)))),
      "suspended_surface_length": selected.get("suspended_surface_length", data.get("suspended_surface_length", selected.get("L_sus", data.get("L_sus", 0.0)))),
      "ground_line_warning_text": selected.get("ground_line_warning_text", data.get("ground_line_warning_text", "")),
      "ground_line_warning": selected.get("ground_line_warning", data.get("ground_line_warning", False)),
      "taut_percentage": selected.get("taut_percentage", data.get("taut_percentage", 0.0)),
      "L_bottom_chain_suspended": selected.get("L_bottom_chain_suspended", data.get("L_bottom_chain_suspended", 0.0)),
      "L_top_chain": selected.get("L_top_chain", data.get("L_top_chain", 0.0)),
      "L_taut": selected.get("L_taut", data.get("L_taut", 0.0)),
      "T_fairlead": selected.get("T_fairlead", data.get("T_fairlead", np.nan)),
      "T_padeye": selected.get("T_padeye", data.get("T_padeye", np.nan)),
      "T_max": selected.get("T_max", data.get("T_max", np.nan)),
      "taut_min_tension": selected.get("taut_min_tension", data.get("taut_min_tension", np.nan)),
      "taut_max_tension": selected.get("taut_max_tension", data.get("taut_max_tension", np.nan)),
      "line_fos_fairlead": selected.get("line_fos_fairlead", data.get("line_fos_fairlead", 0.0)),
      "line_fos_padeye": selected.get("line_fos_padeye", data.get("line_fos_padeye", 0.0)),
      "line_fos_chain": selected.get("line_fos_chain", data.get("line_fos_chain", 0.0)),
      "line_fos_rope": selected.get("line_fos_rope", data.get("line_fos_rope", 0.0)),
      "line_fos": selected.get("line_fos", data.get("line_fos", 0.0)),
      "HF": selected.get("HF", data.get("HF", 0.0)),
      "VF": selected.get("VF", data.get("VF", 0.0)),
      "HA": selected.get("HA", data.get("HA", 0.0)),
      "VA": selected.get("VA", data.get("VA", 0.0)),
      "profile_label": _info_profile_label,
      "bridle_branch_labels": [
          f"T{selected_turbine_index + 1}L{selected_line_index + 1}a",
          f"T{selected_turbine_index + 1}L{selected_line_index + 1}b",
      ] if selected.get("bridle_geometry") else [],
  })

  # Re-apply branch-specific values after the standard information merge so
  # the selected a/b branch is the value shown in the infobox.
  if bridle_navigation and selected_branch_index in (1, 2):
    _branch_anchor_result = (selected.get("bridle_anchor_results") or {}).get(
        f"Branch {selected_branch_index}", {}
    )
    selected_info_data["anchor_results"] = (
        {primary_anc: _branch_anchor_result}
        if _branch_anchor_result else selected_info_data.get("anchor_results", {})
    )
    selected_info_data["T_padeye"] = _branch_load.get(
        "tension", selected_info_data.get("T_padeye", np.nan)
    )
    selected_info_data["HA"] = _branch_load.get(
        "HA", selected_info_data.get("HA", np.nan)
    )
    selected_info_data["VA"] = _branch_load.get(
        "VA", selected_info_data.get("VA", np.nan)
    )
    selected_info_data["profile_label"] = _branch_name
    selected_info_data["bridle_selected_branch"] = (
        "a" if selected_branch_index == 1 else "b"
    )
    if system_type == "Taut" and selected_bridle_anchor is not None:
      # The selected branch—not the former centreline anchor—is authoritative
      # for the a/b profile's local raster depth, padeye/top and anchor bottom.
      # Keeping these together also makes the infobox use the same Z-plane as
      # the marker, anchor body and purple penetration reference.
      _info_padeye_z = selected_taut_bridle_padeye_z
      _info_bottom_z = selected_taut_bridle_anchor_bottom_z
      if np.isfinite(_info_padeye_z):
        _info_updates = {
            "water_depth": -float(_info_padeye_z),
            "raster_z": float(_info_padeye_z),
            "seabed_z": float(_info_padeye_z),
            "padeye_z": float(_info_padeye_z),
            "z_anchor_custom": float(_info_padeye_z),
            "anchor_top_z": float(_info_padeye_z),
        }
        if np.isfinite(_info_bottom_z):
          _info_updates["anchor_bottom_z"] = float(_info_bottom_z)
        selected_info_data.update(_info_updates)

  info = build_mooring_info_lines(
      selected_info_data, profile_label=_info_profile_label
  )
  ax1.text(
      0.03, 0.97, "\n".join(info), transform=ax1.transAxes, fontsize=6.2,
      va="top", bbox=dict(boxstyle="round", facecolor="white", alpha=0.9), zorder=10001
  )

  _profile_display_depth = (
      -float(selected_taut_bridle_padeye_z)
      if system_type == "Taut" and np.isfinite(selected_taut_bridle_padeye_z)
      else water_depth
  )
  # Keep the title compact enough for the dedicated 2-D title band.  Splitting
  # the profile identifier from the anchor/depth detail prevents a long title
  # from reaching into the Previous/Next controls above it.
  _profile_title_detail = str(primary_anc)
  if np.isfinite(_profile_display_depth):
    _profile_title_detail += f" | {_profile_display_depth:.1f} m depth"
  ax1.set_title(
      f"Mooring Line Profile — {_info_profile_label}\n{_profile_title_detail}",
      fontweight="bold", fontsize=10.5, y=1.0, pad=7,
  )
  ax1.set_xlabel("Horizontal Distance [m]")
  ax1.set_ylabel("Elevation [m]")
  ax1.grid(True, linestyle=":", alpha=0.6)
  # Tell the shared legend cleaner that this is the 2-D Semi-Taut profile.
  # Semi-Taut needs separate Bottom Chain / Middle Rope / Top Chain entries.
  ax1._semi_taut_2d_legend = (str(system_type).strip().lower() == "semi-taut")
  _dedupe_mooring_legend(ax1)
  # Robust horizontal framing for short and very long lines.  The previous
  # fixed ``xf + 25`` framing assumed that the fairlead was the largest x-value
  # ever rendered.  That is not always true for terrain/bridle overlays, and it
  # can make long profiles look clipped or stretched at the ends.  Frame from
  # the actual displayed geometry instead; this is a rendering-only change.
  _x_candidates = [
      np.asarray(x_plot, dtype=float),
      np.asarray(sub_x, dtype=float),
  ]
  if rope_nodes is not None:
      try:
          _x_candidates.append(np.asarray([rope_nodes[0], rope_nodes[2]], dtype=float))
      except Exception:
          pass
  if selected_ground is not None:
      try:
          _x_candidates.append(np.asarray(selected_ground.get("x", []), dtype=float))
      except Exception:
          pass
  if bridle_2d and bridle_2d.get("enabled"):
      for _bp_frame in bridle_2d.get("profiles", []):
          try:
              _bx_frame = _bridle_branch_2d_x(_bp_frame)
              if _bx_frame.size:
                  _x_candidates.append(np.asarray(_bx_frame, dtype=float))
          except Exception:
              pass
  _x_candidates.append(np.asarray([0.0, float(xf)], dtype=float))
  _x_all = np.concatenate([a[np.isfinite(a)] for a in _x_candidates if a.size and np.any(np.isfinite(a))])
  if _x_all.size:
      _xmin_render = float(np.min(_x_all))
      _xmax_render = float(np.max(_x_all))
  else:
      _xmin_render, _xmax_render = 0.0, max(float(xf), 25.0)
  _x_span_render = max(_xmax_render - _xmin_render, 1.0)
  _x_margin_render = max(15.0, 0.04 * _x_span_render)
  ax1.set_xlim(_xmin_render - _x_margin_render, _xmax_render + _x_margin_render)

  # -------------------------------------------------------------------------
  # 3D FARM — ORIGINAL anchor/padeye/mooring rendering, raster seabed surface.
  # -------------------------------------------------------------------------
  min_spacing = (2.0 * anchor_radius) + buffer_zone
  rows, cols = get_optimal_grid_dimensions(num_turbines)
  headings = generate_mooring_headings(
      num_lines,
      triad=bool(data.get("triad", False)),
      inter_arm_angle_deg=float(data.get("inter_arm_angle", 0.0)),
      angle_between_mooring_arms_deg=float(
          data.get("angle_between_mooring_arms", 360.0 / max(num_lines, 1))
      ),
  )
  base_circle_angles = np.linspace(0, 2 * np.pi, 360)
  circle_angles = np.sort(np.unique(np.concatenate((base_circle_angles, np.radians(headings)))))

  if turbine_coords:
    xs = np.asarray([p[0] for p in turbine_coords], dtype=float)
    ys = np.asarray([p[1] for p in turbine_coords], dtype=float)
    margin = max(anchor_radius * 1.2, 250.0)
    gx = np.linspace(xs.min() - margin, xs.max() + margin, 45)
    gy = np.linspace(ys.min() - margin, ys.max() + margin, 45)
    GX, GY = np.meshgrid(gx, gy)
    GZ = np.asarray(render_raster_sampler.sample(GX, GY), dtype=float)
    if np.any(np.isfinite(GZ)):
      ax2.plot_surface(GX, GY, np.where(np.isfinite(GZ), GZ, np.nan),
                       cmap=NAVIA_CMAP, linewidth=0, antialiased=True, alpha=0.78,
                       rcount=45, ccount=45, zorder=0)

  r_anchor = anchor_radius
  depth_failure_legend_shown = False
  anchor_bottom_z_values = []
  for i, (cx, cy) in enumerate(turbine_coords):
    turbine_lines = line_results[i]
    valid_depths = [v.get("local_depth", np.nan) for v in turbine_lines if np.isfinite(v.get("local_depth", np.nan))]

    # ORIGINAL turbine and guide-circle rendering.
    fairlead_z = fairlead_draft
    ax2.scatter(cx, cy, fairlead_z, color="blue", s=40, marker="^",
                edgecolor="black", linewidths=0.8, alpha=_asset_alpha(), zorder=44,
                label="Floating Turbine" if i == 0 else "")
    ax2.text(
        cx, cy, fairlead_z + 15.0, f"T{i+1}",
        color=UI_NAVY, fontsize=8, fontweight="bold",
        ha="center", va="bottom", alpha=_asset_alpha(), zorder=45,
    )

    # There is one anchor-radius guide per turbine, not one per mooring line.
    # Previously this was drawn inside the line loop below using each anchor's
    # local seabed depth.  That produced several nearly coincident radii in the
    # unified 3-D view.  Sampling the raster along the ring makes this single
    # guide follow the displayed seabed and keeps it geometrically meaningful.
    _ring_x = cx + r_anchor * np.cos(circle_angles)
    _ring_y = cy + r_anchor * np.sin(circle_angles)
    _ring_fallback_z = (
        -float(np.nanmedian(valid_depths))
        if valid_depths else -float(water_depth)
    )
    if not np.isfinite(_ring_fallback_z):
      _ring_fallback_z = 0.0
    _ring_z = np.full_like(circle_angles, _ring_fallback_z, dtype=float)
    if raster_path:
      try:
        _sampled_ring_z = np.asarray(
            render_raster_sampler.sample(_ring_x, _ring_y), dtype=float
        )
        if _sampled_ring_z.shape == _ring_z.shape:
          _ring_z = np.where(
              np.isfinite(_sampled_ring_z), _sampled_ring_z, _ring_fallback_z
          )
      except Exception:
        # The guide remains usable at the local median depth if the raster
        # cannot be sampled; the mooring calculations themselves are unchanged.
        pass
    ax2.plot(
        _ring_x, _ring_y, _ring_z, "r--", lw=0.8, alpha=0.6,
        zorder=5, label="Anchor Radius" if i == 0 else "",
    )

    # Render each individual mooring line after the one shared guide ring.
    for j, angle in enumerate(headings):
      local = turbine_lines[j]
      local_depth = local.get("local_depth", np.nan)
      seabed_z = (
          -float(local_depth)
          if np.isfinite(local_depth) else float(_ring_fallback_z)
      )

      if not local.get("success"):
        continue

      # Excel H30 is a siting gate, not a reason to discard an otherwise
      # solvable line.  Keep every geometry segment visible and render the
      # affected TnLn red.  True solver failures above remain intentionally
      # excluded because there is no valid geometry to display.
      line_depth_failed = not bool(local.get("_raster_depth_band_pass", True))
      line_chain_colour = UI_FAIL if line_depth_failed else "black"
      line_rope_colour = UI_FAIL if line_depth_failed else "deepskyblue"
      line_rope_outline_colour = "#7a1010" if line_depth_failed else "black"
      if line_depth_failed and not depth_failure_legend_shown:
        ax2.plot(
            [], [], [], color=UI_FAIL, lw=2.6,
            label="Depth-threshold failed mooring",
        )
        depth_failure_legend_shown = True

      # Local result contains the ORIGINAL solver geometry at this local depth.
      lx = np.asarray(local.get("x_plot", []), dtype=float)
      lz = np.asarray(local.get("z_plot", []), dtype=float)
      lsub_x = local.get("sub_x")
      lsub_z = local.get("sub_z")
      lrope_nodes = local.get("rope_nodes")
      lX_td = local.get("X_td")
      if (
          system_type != "Taut"
          and i == selected_turbine_index
          and j == selected_line_index
          and selected_ground is not None
      ):
        local_ground = selected_ground
      else:
        local_ground = terrain_following_line_geometry(
            data,
            local,
            raster_path,
            (cx, cy),
            float(angle),
            raster_sampler=render_raster_sampler,
        ) if system_type != "Taut" else None

      # Build the inverted bridle for EVERY turbine/line, not just the selected
      # profile. This makes the complete farm genuinely navigable and ensures
      # every TnLn profile has its own two branch anchors/lines.
      local_bridle = None
      if bool(data.get("inverted_bridle", False)) and local.get("success"):
        try:
          local["inverted_bridle"] = True
          local["triad"] = bool(data.get("triad", False))
          local["upper_equals_lower"] = bool(data.get("upper_equals_lower", True))
          local["lower_joint_pos"] = float(data.get("lower_joint_pos", 0.5))
          local["inter_anchor_angle"] = float(data.get("inter_anchor_angle", 30.0))
          local_bridle = prepare_selected_inverted_bridle(
              local,
              float(angle),
              (cx, cy),
              raster_path=raster_path,
              selected_ground=local_ground,
              raster_sampler=render_raster_sampler,
          )
        except Exception as _local_bridle_exc:
          local["bridle_geometry_error"] = str(_local_bridle_exc)
          local_bridle = None

      # The original radial padeye is retained for the centreline trunk.  An
      # inverted bridle, however, has two *physical* padeyes at its branch
      # anchor positions; those are what must be shown in the 3-D dashboard.
      padeye_3d_x = cx + r_padeye * np.cos(np.radians(angle))
      padeye_3d_y = cy + r_padeye * np.sin(np.radians(angle))
      ang = np.radians(angle)
      ux, uy = np.cos(ang), np.sin(ang)
      anc_type_upper = primary_anc.upper()

      # Use every solver-produced branch anchor when a bridle is active.  The
      # old renderer drew a phantom centreline padeye and only rendered the
      # real branch padeyes for T1L1, which made other turbines/lines appear
      # to be missing anchors despite their solved geometry being valid.
      _padeye_specs = []
      if local_bridle and local_bridle.get("enabled"):
        for _branch_anchor in local_bridle.get("anchors", []):
          try:
            _padeye_specs.append((
                float(_branch_anchor["x"]),
                float(_branch_anchor["y"]),
                float(_branch_anchor["z"]),
                float(_branch_anchor["heading_deg"]),
            ))
          except (KeyError, TypeError, ValueError):
            continue

      if not _padeye_specs:
        if system_type == "Taut":
          # Taut anchor head/padeye is at the seabed surface. Its physical
          # length then extends downward to the actual maximum penetration.
          z_anchor_3d = float(seabed_z)
        elif any(k in anc_type_upper for k in ("SUCTION", "DRIVEN", "DRILLED")):
          frac = float(
              (data.get("padeye_params") or {}).get(
                  primary_anc, {}
              ).get("position_fraction", 0.50)
          )
          _, _, z_anchor_3d = get_pile_anchor_vertical_geometry(
              seabed_z * -1.0, anchor_depth, anchor_height,
              padeye_fraction=frac,
          )
        else:
          z_anchor_3d = (
              lsub_z[0] if lsub_z is not None and len(lsub_z) > 0
              else seabed_z - anchor_depth + 0.5 * anchor_height
          )
        _padeye_specs.append((
            float(padeye_3d_x),
            float(padeye_3d_y),
            float(z_anchor_3d),
            float(angle),
        ))

      if system_type == "Taut":
        # Include actual physical anchor bottoms in the view limits.  A
        # branch pile can be much longer than the generic seabed margin and
        # otherwise looks truncated even though it has been drawn correctly.
        for _, _, _padeye_z_for_limit, _ in _padeye_specs:
          _anchor_bottom_for_limit = (
              float(_padeye_z_for_limit) - float(anchor_height)
          )
          if np.isfinite(_anchor_bottom_for_limit):
            anchor_bottom_z_values.append(_anchor_bottom_for_limit)

      for _padeye_idx, (
          _padeye_x, _padeye_y, _padeye_z, _padeye_heading
      ) in enumerate(_padeye_specs):
        _draw_dashboard_anchor_3d(
            ax2,
            _padeye_x,
            _padeye_y,
            _padeye_z,
            _padeye_heading,
            system_type,
            primary_anc,
            anchor_width,
            anchor_height,
            anchor_depth,
            seabed_z,
            data,
            show_legend=(i == 0 and j == 0 and _padeye_idx == 0),
            opacity=turbine_opacity,
        )

      # 3D mooring-line rendering.
      #
      # With an inverted bridle active, the original line is physically split:
      #   Fairlead -> Lower Joint
      #   Lower Joint -> Anchor A (T1Lna)
      #   Lower Joint -> Anchor B (T1Lnb)
      #
      # The original lower/grounded section is therefore NOT plotted as a
      # third line.  Each branch has its own plan direction and raster-ground
      # interaction.
      ground_display_offset = 0.75
      if local_bridle and local_bridle.get("enabled"):
        _jx, _jy, _jz = local_bridle["joint"]

        # Plot only the original trunk above Joint 1.
        if lrope_nodes is not None:
          x1_n, z1_n, x2_n, z2_n = lrope_nodes
          rx = np.asarray(lx, dtype=float)
          rz = np.asarray(lz, dtype=float)
          _tm = rx >= x1_n - 1e-9
          if np.any(_tm):
            _tx = np.concatenate(([x1_n], rx[_tm]))
            _tz = np.concatenate(([z1_n], rz[_tm]))
            _keep = np.r_[True, np.diff(_tx) > 1e-10]
            _tx, _tz = _tx[_keep], _tz[_keep]
            _px3 = cx + (r_padeye - _tx[::2]) * ux
            _py3 = cy + (r_padeye - _tx[::2]) * uy
            _pz3 = _tz[::2]
            if system_type == "Taut":
              ax2.plot(_px3, _py3, _pz3, color=line_rope_outline_colour, alpha=_asset_alpha(0.98), lw=4.2, zorder=30)
            ax2.plot(
                _px3, _py3, _pz3,
                color=line_rope_colour, alpha=_asset_alpha(0.98), lw=2.6, zorder=31,
                label="Synthetic Rope" if (i == 0 and j == 0) else "",
            )
            # Top chain remains part of the trunk.
            _top_m = rx >= x2_n - 1e-9
            if np.any(_top_m):
              _top_x = np.concatenate(([x2_n], rx[_top_m], [float(local.get("xf", rx[-1]))]))
              _top_z = np.concatenate(([z2_n], rz[_top_m], [float(local.get("fairlead_draft", fairlead_draft))]))
              _keep = np.r_[True, np.diff(_top_x) > 1e-10]
              _top_x, _top_z = _top_x[_keep], _top_z[_keep]
              ax2.plot(
                  cx + (r_padeye - _top_x[::2]) * ux,
                  cy + (r_padeye - _top_x[::2]) * uy,
                  _top_z[::2],
                  color=line_chain_colour, alpha=_asset_alpha(0.95), lw=2.2, zorder=30,
                  label="Top Chain" if (i == 0 and j == 0) else "",
              )
        else:
          if local_ground is not None:
            _full_x = np.asarray(local_ground["x"], dtype=float)
            _full_z = np.asarray(local_ground["z"], dtype=float)
          else:
            _full_x = np.concatenate((lsub_x, lx[1:])) if lsub_x is not None else lx
            _full_z = np.concatenate((lsub_z, lz[1:])) if lsub_z is not None else lz

          _tm = _full_x >= float(local_bridle["joint_local"][0]) - 1e-9
          _tx = np.concatenate(([float(local_bridle["joint_local"][0])], _full_x[_tm]))
          _tz = np.concatenate(([float(local_bridle["joint_local"][1])], _full_z[_tm]))
          _keep = np.r_[True, np.diff(_tx) > 1e-10]
          _tx, _tz = _tx[_keep], _tz[_keep]
          _px3 = cx + (r_padeye - _tx[::4]) * ux
          _py3 = cy + (r_padeye - _tx[::4]) * uy
          _pz3 = _tz[::4]
          if system_type == "Taut":
            ax2.plot(_px3, _py3, _pz3, color=line_rope_outline_colour, alpha=_asset_alpha(0.95), lw=4.0, zorder=30)
            ax2.plot(_px3, _py3, _pz3, color=line_rope_colour, alpha=_asset_alpha(0.98), lw=2.4, zorder=31,
                     label=f"{system_type} Trunk" if (i == 0 and j == 0) else "")
          else:
            ax2.plot(_px3, _py3, _pz3, color=line_chain_colour, alpha=_asset_alpha(0.95), lw=2.2, zorder=30,
                     label=f"{system_type} Trunk" if (i == 0 and j == 0) else "")

        # Both physical branches, including raster-resting groundline.
        for _bi, _branch in enumerate(local_bridle.get("profiles", []), start=1):
          _suffix = "a" if _bi == 1 else "b"
          _line_name = f"T{i+1}L{j+1}{_suffix}"
          _bx = np.asarray(_branch.get("x", []), dtype=float)
          _by = np.asarray(_branch.get("y", []), dtype=float)
          _bz = np.asarray(_branch.get("z", []), dtype=float).copy()
          if _bx.size < 2 or _by.size != _bx.size or _bz.size != _bx.size:
            continue

          # Lift only raster-contact portions above the 3-D raster surface.
          _ground_mask = np.asarray(
              _branch.get("ground_mask", np.zeros(_bx.shape, dtype=bool)),
              dtype=bool,
          )
          _bz[_ground_mask] += ground_display_offset

          _show_label = (
              _line_name
              if (i == 0 and j == 0 and (not bridle_navigation or _bi == selected_branch_index))
              else ""
          )
          if system_type == "Taut":
            ax2.plot(_bx, _by, _bz, color=line_rope_outline_colour, alpha=_asset_alpha(0.98), lw=4.2, zorder=30)
            ax2.plot(_bx, _by, _bz, color=line_rope_colour, alpha=_asset_alpha(0.98), lw=2.6, zorder=31,
                     label=_show_label)
          else:
            ax2.plot(_bx, _by, _bz, color=line_chain_colour, alpha=_asset_alpha(0.98), lw=2.4, zorder=31,
                     label=_show_label)

          # TDP and DDP are physical points on EACH bridle branch.
          # Use the branch's own XY curve rather than the original centreline.
          if lX_td is not None and _bx.size >= 2:
            _x_local = float(lX_td)
            _td_idx = int(np.argmin(np.abs(np.asarray(_bridle_branch_2d_x(_branch), dtype=float) - _x_local)))
            _td_x3 = float(_bx[_td_idx]); _td_y3 = float(_by[_td_idx]); _td_z3 = float(_bz[_td_idx])
            ax2.scatter(
                [_td_x3], [_td_y3], [_td_z3], color="yellow", marker="o", s=18,
                edgecolors="black", linewidths=0.5, alpha=_asset_alpha(), zorder=36,
                label=("Touchdown Point (TDP)" if (i == 0 and j == 0 and _bi == selected_branch_index) else ""),
            )

          if lsub_x is not None and len(lsub_x) > 0:
            _ddp_x_local = float(np.asarray(lsub_x, dtype=float)[-1])
            _pfx = np.asarray(_bridle_branch_2d_x(_branch), dtype=float)
            if _pfx.size >= 2:
              _ddp_idx = int(np.argmin(np.abs(_pfx - _ddp_x_local)))
              _ddp_x3 = float(_bx[_ddp_idx]); _ddp_y3 = float(_by[_ddp_idx]); _ddp_z3 = float(_bz[_ddp_idx])
              ax2.scatter(
                  [_ddp_x3], [_ddp_y3], [_ddp_z3], color="yellow", marker="s", s=18,
                  edgecolors="black", linewidths=0.5, alpha=_asset_alpha(), zorder=36,
                  label=("Down Dip Point (DDP)" if (i == 0 and j == 0 and _bi == selected_branch_index) else ""),
              )

        # Semi-Taut already has the Bottom Chain / Rope Joint marker.
        # Avoid drawing a second generic Lower Joint symbol.
        if system_type != "Semi-Taut":
          ax2.scatter(
              [_jx], [_jy], [_jz],
              color="magenta", marker="D", s=28,
              edgecolors="black", linewidths=0.7, alpha=_asset_alpha(), zorder=36,
              label="Lower Joint" if (i == 0 and j == 0) else "",
          )
      elif lrope_nodes is not None:
        x1_n, z1_n, x2_n, z2_n = lrope_nodes
        if local_ground is not None:
          profile_x = np.asarray(local_ground["x"], dtype=float)
          profile_z = np.asarray(local_ground["z"], dtype=float)
          mask_bot = profile_x <= x1_n + 1e-9
          bot_x = profile_x[mask_bot]
          bot_z = profile_z[mask_bot].copy()
          if local_ground.get("x_td") is not None:
            ground_mask = bot_x >= float(local_ground["x_ddp"]) - 1e-9
            bot_z[ground_mask] += ground_display_offset
        else:
          mask_bot = lx <= x1_n
          bot_x = (np.concatenate((lsub_x, lx[mask_bot][1:] if np.any(mask_bot) else []))
                   if lsub_x is not None else lx[mask_bot])
          bot_z = (np.concatenate((lsub_z, lz[mask_bot][1:] if np.any(mask_bot) else []))
                   if lsub_z is not None else lz[mask_bot])
        if len(bot_x) == 0 or bot_x[-1] != x1_n:
          bot_x = np.append(bot_x, x1_n)
          bot_z = np.append(bot_z, z2_n if len(bot_z) == 0 else bot_z[-1])
        ax2.plot(cx + (r_padeye - bot_x[::2]) * ux,
                 cy + (r_padeye - bot_x[::2]) * uy,
                 bot_z[::2], color=line_chain_colour, alpha=_asset_alpha(0.95), lw=2.2, zorder=30,
                 label="Grounded Mooring Line" if (i == 0 and j == 0) else "")
        ax2.plot(cx + (r_padeye - np.array([x1_n, x2_n])) * ux,
                 cy + (r_padeye - np.array([x1_n, x2_n])) * uy,
                 [z1_n, z2_n], color=line_rope_colour, alpha=_asset_alpha(0.98), lw=2.6, zorder=31,
                 label="Synthetic Rope" if (i == 0 and j == 0) else "")
      else:
        if local_ground is not None:
          full_x = np.asarray(local_ground["x"], dtype=float)
          full_z = np.asarray(local_ground["z"], dtype=float).copy()
          ground_mask = (full_x >= float(local_ground["x_ddp"]) - 1e-9) & (
              full_x <= float(local_ground["x_td"]) + 1e-9
          )
          full_z[ground_mask] += ground_display_offset
        else:
          full_x = np.concatenate((lsub_x, lx[1:])) if lsub_x is not None else lx
          full_z = np.concatenate((lsub_z, lz[1:])) if lsub_z is not None else lz
        line_color = line_rope_colour if system_type == "Taut" else line_chain_colour
        ax2.plot(cx + (r_padeye - full_x[::4]) * ux,
                 cy + (r_padeye - full_x[::4]) * uy,
                 full_z[::4],
                 color=line_color, alpha=_asset_alpha(0.95), lw=2.2, zorder=30,
                 label="Mooring Line" if (i == 0 and j == 0) else "")


      if lX_td is not None and not (local_bridle and local_bridle.get("enabled")):
        r_tdp = r_padeye - float(lX_td)
        td_x = cx + r_tdp * ux
        td_y = cy + r_tdp * uy

        line_x_td = np.asarray(lx, dtype=float)
        line_z_td = np.asarray(lz, dtype=float)
        if system_type != "Taut" and lsub_x is not None and lsub_z is not None:
          sx = np.asarray(lsub_x, dtype=float)
          sz = np.asarray(lsub_z, dtype=float)
          if sx.size and sz.size:
            line_x_td = np.concatenate((sx, line_x_td[1:]))
            line_z_td = np.concatenate((sz, line_z_td[1:]))

        if local_ground is not None:
          xq = float(np.clip(float(lX_td), local_ground["x"][0], local_ground["x"][-1]))
          td_z = float(np.interp(xq, local_ground["x"], local_ground["z"]))
        else:
          finite = np.isfinite(line_x_td) & np.isfinite(line_z_td)
          if np.count_nonzero(finite) >= 2:
            lx_td = line_x_td[finite]
            lz_td = line_z_td[finite]
            order = np.argsort(lx_td)
            lx_td = lx_td[order]
            lz_td = lz_td[order]
            xq = float(np.clip(float(lX_td), lx_td[0], lx_td[-1]))
            td_z = float(np.interp(xq, lx_td, lz_td))
          else:
            td_z = float(seabed_z)

        ax2.scatter([td_x], [td_y], [td_z], color="yellow", s=18,
                    edgecolors="black", linewidths=0.5, alpha=_asset_alpha(), zorder=36,
                    label="Touchdown Point (TDP)" if i == 0 and j == 0 else "")

      if lsub_x is not None and lsub_z is not None and len(lsub_x) > 0 and not (local_bridle and local_bridle.get("enabled")):
        r_ddp = r_padeye - lsub_x[-1]
        ddp_x = cx + r_ddp * ux
        ddp_y = cy + r_ddp * uy
        if local_ground is not None:
          ddp_z = float(np.interp(
              local_ground["x_ddp"], local_ground["x"], local_ground["z"]
          ))
        else:
          ddp_z = float(lsub_z[-1])
        ax2.scatter([ddp_x], [ddp_y], [ddp_z], color="yellow", marker="s", s=18,
                    edgecolors="black", linewidths=0.5, alpha=_asset_alpha(), zorder=36,
                    label="Down Dip Point (DDP)" if i == 0 and j == 0 else "")

  # Farm perimeter follows the raster seabed, preserving original purple style.
  # Match the exported survey perimeter: anchor radius + additional 50 m survey allowance.
  if turbine_coords:
    min_cx, max_cx = min(p[0] for p in turbine_coords), max(p[0] for p in turbine_coords)
    min_cy, max_cy = min(p[1] for p in turbine_coords), max(p[1] for p in turbine_coords)
    survey_perimeter_buffer_m = 50.0
    buffer = anchor_radius + survey_perimeter_buffer_m
    xmin, xmax, ymin, ymax = min_cx-buffer, max_cx+buffer, min_cy-buffer, max_cy+buffer
    num_p = 30
    edge_x = np.concatenate([np.linspace(xmin, xmax, num_p), np.full(num_p, xmax),
                             np.linspace(xmax, xmin, num_p), np.full(num_p, xmin)])
    edge_y = np.concatenate([np.full(num_p, ymin), np.linspace(ymin, ymax, num_p),
                             np.full(num_p, ymax), np.linspace(ymax, ymin, num_p)])
    edge_z = np.asarray(
        render_raster_sampler.sample(edge_x, edge_y), dtype=float
    )
    ax2.plot(edge_x, edge_y, edge_z, color="purple", linestyle="-", linewidth=1.5,
             alpha=0.7, zorder=5, label="Farm Perimeter")

  span = max(total_width_x, total_height_y) + (2.0 * (anchor_radius + 50.0)) + 5000.0
  ax2.set_xlim(origin_x - span / 2, origin_x + span / 2)
  ax2.set_ylim(origin_y - span / 2, origin_y + span / 2)
  finite_z = []
  for lines in line_results:
    for local in lines:
      if np.isfinite(local.get("local_depth", np.nan)):
        finite_z.append(-float(local["local_depth"]))
      if local.get("success"):
        zz = np.asarray(local.get("z_plot", []), dtype=float)
        finite_z.extend(zz[np.isfinite(zz)].tolist())
  finite_z.extend(anchor_bottom_z_values)
  if finite_z:
    zmin = min(finite_z) - 35.0
    zmax = max(50.0, fairlead_draft + 10.0)
    ax2.set_zlim(zmin, zmax)
  else:
    ax2.set_zlim(-water_depth - 35, 50.0)
  ax2.set_box_aspect([1, 1, 0.5])
  ax2.set_title(f"Farm Layout ({primary_anc})", fontweight="bold")
  _dedupe_mooring_legend(ax2)

  render_raster_sampler.close()
  return [selected]


def infer_failure_stage(system, data):
  """Classify the earliest failed/pre-feasibility stage without inventing downstream results.

  Stages are intentionally ordered so the Failure Criteria report can distinguish:
    geometry -> ground/angle -> tension/line capacity -> anchor suitability.

  A successful result is considered fully evaluated.  For failed results the
  message is the governing source because the solver returns immediately at the
  first failed check.
  """
  if not isinstance(data, dict):
    return "unknown"
  if bool(data.get("success")):
    return "complete"

  msg = str(data.get("msg", "") or "").strip().lower()
  sys = str(system or data.get("system_type", "")).strip().lower()

  geometry_terms = (
      "target span exceeds line length",
      "no valid generated taut length",
      "invalid net suspended span",
      "no suspended line remains",
      "semi-taut solver failed to converge",
      "solution has no suspended section",
      "geometry",
      "connect to the fairlead",
      "reverses horizontally",
      "broken/kinked",
      "non-finite coordinates",
      "zero-length segments",
      "invalid net",
  )
  ground_terms = (
      "insufficient ground line",
      "excessive ground line",
      "ground line",
      "grounded line",
  )
  tension_terms = (
      "taut section is slack",
      "taut line is under-tensioned",
      "under-tensioned",
      "tension profile",
      "mbl/fos screening failed",
      "tension exceeds rope mbl",
      "maximum tension",
      "tension exceeds",
  )
  anchor_terms = (
      "dea uplift failure",
      "geotechnical capacity failed",
      "anchor suitability",
  )

  # Numerical geometry takes absolute precedence over any stale/downstream
  # values that may still exist in a failed result dictionary.  A line that is
  # shorter than the calculated minimum geometric length is an upstream
  # geometry failure, so nothing downstream may be reported as evaluated.
  actual_line = _safe_float(data.get("line_length", data.get("line_length_main", np.nan)))
  min_geom = _safe_float(data.get("min_geometric_length", np.nan))
  if np.isfinite(actual_line) and np.isfinite(min_geom) and actual_line + 1e-6 < min_geom:
    return "geometry"

  if any(t in msg for t in geometry_terms):
    return "geometry"
  if sys != "taut" and any(t in msg for t in ground_terms):
    return "ground"
  if any(t in msg for t in anchor_terms):
    return "anchor"
  if any(t in msg for t in tension_terms):
    return "tension"

  # Conservative fallback: only claim a later stage if the solver explicitly
  # returned the numerical inputs required by that stage.
  if all(k in data and np.isfinite(_safe_float(data.get(k))) for k in ("HA", "VA", "HF", "VF")):
    return "tension"
  if sys != "taut" and "ground_ratio" in data:
    return "ground"
  if "min_geometric_length" in data or "excess_line" in data:
    return "geometry"
  return "unknown"


def _safe_float(value):
  try:
    v = float(value)
    return v
  except Exception:
    return np.nan


def failure_evaluation_flags(system, data):
  """Return the report-stage dependency state for pre-feasibility reporting."""
  stage = infer_failure_stage(system, data)
  is_taut = str(system).strip().lower() == "taut"
  complete = stage == "complete"

  # Geometry is evaluated when the report can establish actual vs minimum.
  geometry_eval = (
      np.isfinite(_safe_float(data.get("line_length", data.get("line_length_main", np.nan))))
      and np.isfinite(_safe_float(data.get("min_geometric_length", np.nan)))
  ) or (stage == "geometry")

  # Ground/angle only exist for Catenary/Semi-Taut. A geometry failure prevents
  # both from being valid downstream calculations.
  ground_eval = (not is_taut) and stage not in ("geometry", "unknown") and (
      "ground_ratio" in data or "grounded_surface_length" in data or "L_bot" in data
  )
  angle_eval = (not is_taut) and stage not in ("geometry", "unknown") and (
      "suspended_line_angle_deg" in data
  )

  # Pre-feasibility dependencies are criterion-specific.  An anchor failure
  # must not suppress line tension, MBL or line FoS results that were already
  # calculated.  Geometry/ground failures, however, block downstream tension
  # criteria because no valid downstream state exists.
  upstream_tension_block = stage in ("geometry", "ground", "unknown")
  tension_eval = (not upstream_tension_block) and all(
      np.isfinite(_safe_float(data.get(k, np.nan)))
      for k in ("HA", "VA", "HF", "VF")
  )

  max_tension_value = _safe_float(data.get("T_max", np.nan))
  if not np.isfinite(max_tension_value):
    max_tension_value = _safe_float(data.get("taut_max_tension", np.nan))

  break_eval = (not upstream_tension_block) and np.isfinite(max_tension_value)

  # Line FoS is independent of anchor FoS.  A line FoS explicitly returned by
  # the solver remains valid even when the eventual failure is anchor-related.
  line_fos_eval = (not upstream_tension_block) and any(
      np.isfinite(_safe_float(data.get(k, np.nan))) for k in (
          "line_fos", "line_fos_chain", "line_fos_rope", "line_fos_fairlead", "line_fos_padeye"
      )
  )

  # Anchor FoS is evaluated only when the anchor-capacity routine actually
  # supplied a finite FoS.  An anchor uplift/capacity failure can therefore
  # coexist with an independently evaluated line FoS.
  anchor_fos_eval = (stage not in ("geometry", "ground", "unknown")) and any(
      isinstance(v, dict) and np.isfinite(_safe_float(v.get("fos", np.nan)))
      for v in (data.get("anchor_results") or {}).values()
  )

  tension_components_eval = tension_eval and all(
      np.isfinite(_safe_float(data.get(k, np.nan))) for k in ("HA", "VA", "HF", "VF")
  )

  return {
      "stage": stage,
      "geometry": geometry_eval,
      "ground": ground_eval,
      "angle": angle_eval,
      "tension": tension_eval,
      "break": break_eval,
      "line_fos": line_fos_eval,
      "anchor_fos": anchor_fos_eval,
      "tension_components": tension_components_eval,
      "taut_ground_applicable": not is_taut,
  }

def enrich_failure_report_data(data, system, anchor_db):
  """Add only non-downstream metadata needed to explain a failed case.

  IMPORTANT: This function deliberately does *not* reconstruct H/V tensions,
  MBL/FoS, grounded ratios, or anchor capacity after an upstream solver failure.
  Those quantities must come from the actual solver stage that calculated them.
  This prevents the report from displaying plausible-looking downstream values
  for an invalid/unevaluated configuration.
  """
  d = dict(data or {})
  if d.get("success", False):
    d.setdefault("_failure_stage", "complete")
    return d

  stage = infer_failure_stage(system, d)
  d["_failure_stage"] = stage

  # Preserve the failure reason exactly; only add static geometry metadata that
  # does not depend on a successfully solved downstream mooring state.
  anc = d.get("primary_anc") or d.get("anchor")
  try:
    anc_w, anc_h, anc_depth = get_anchor_dimensions(anchor_db, anc)
  except Exception:
    anc_w = _safe_float(d.get("anchor_width", np.nan))
    anc_h = _safe_float(d.get("anchor_height", np.nan))
    anc_depth = _safe_float(d.get("anchor_depth", np.nan))

  water_depth = _safe_float(d.get("water_depth", np.nan))
  anchor_radius = _safe_float(d.get("anchor_radius", np.nan))
  fairlead_draft = _safe_float(d.get("fairlead_draft", np.nan))
  fairlead_radius = _safe_float(d.get("fairlead_radius", np.nan))
  if not all(np.isfinite(v) for v in (water_depth, anchor_radius, fairlead_draft, fairlead_radius)):
    return d

  padeye_params = d.get("padeye_params")
  try:
    _sub_x, _sub_z, L_sub, x_sub_span, z_anchor, padeye_angle = get_subsurface_geometry(
        water_depth, anchor_db, anc, padeye_params=padeye_params
    )
  except Exception:
    L_sub = x_sub_span = z_anchor = padeye_angle = np.nan

  if system in ("Catenary", "Semi-Taut") and np.isfinite(L_sub) and np.isfinite(x_sub_span):
    r_padeye = anchor_radius - float(anc_w)
    xf = abs(r_padeye - fairlead_radius)
    zf = fairlead_draft + water_depth
    xf_net = xf - x_sub_span
    if np.isfinite(xf_net):
      minimum = float(L_sub + np.hypot(max(xf_net, 0.0), zf))
      d.setdefault("min_geometric_length", minimum)
      d.setdefault("min_geometric_main_length", float(np.hypot(max(xf_net, 0.0), zf)))
      actual = _safe_float(d.get("line_length", np.nan))
      if np.isfinite(actual):
        d.setdefault("excess_line", actual - minimum)
      d.setdefault("padeye_angle_deg", padeye_angle)
      d.setdefault("padeye_z", z_anchor)
      d.setdefault("anchor_width", anc_w)
      d.setdefault("anchor_height", anc_h)
      d.setdefault("anchor_depth", anc_depth)

  elif system == "Taut":
    r_padeye = anchor_radius - (float(anc_w) / 2.0)
    # Preserve the authoritative solver/raster padeye elevation.  Never move
    # the Taut padeye based on nominal penetration depth.
    z_anchor_taut = _safe_float(d.get("padeye_z", d.get("seabed_z", -water_depth)))
    if not np.isfinite(z_anchor_taut):
      z_anchor_taut = -float(water_depth)
    xf = abs(r_padeye - fairlead_radius)
    zf = fairlead_draft - z_anchor_taut
    direct_span = float(np.hypot(xf, zf))
    d.setdefault("direct_geometric_span", direct_span)
    # The generated taut length is the actual design value only when present in
    # the solver result. Do not overwrite it with a fabricated downstream value.
    d.setdefault("z_anchor_custom", z_anchor_taut)
    d.setdefault("anchor_width", anc_w)
    d.setdefault("anchor_height", anc_h)
    d.setdefault("anchor_depth", anc_depth)
    d.setdefault(
        "padeye_angle_deg",
        (padeye_params or {}).get(anc, {}).get("angle_deg", np.nan)
        if isinstance(padeye_params, dict) else np.nan,
    )

  return d


def show_failure_criteria_window(parent_fig=None, data_provider=None):
  """Open/reuse a live, report-style failure-criteria window.

  The report is modeless and refreshes from the active dashboard configuration.
  It is intentionally report-like: the styling is stable, while values and
  failure state update whenever the selected system/anchor/profile changes.
  """
  try:
    import tkinter as tk
  except Exception as exc:
    print(f"[WARNING] Could not import tkinter for failure report: {exc}")
    return None

  if parent_fig is not None:
    existing = getattr(parent_fig, "_failure_criteria_report", None)
    if existing and existing.get("window") is not None:
      try:
        existing["window"].deiconify()
        existing["window"].lift()
        existing["window"].focus_force()
        existing["refresh"]()
        return existing["refresh"]
      except Exception:
        try:
          parent_fig._failure_criteria_report = None
        except Exception:
          pass

  win = tk.Toplevel()
  win.title("Mooring Failure Criteria — Live Configuration Report")
  try:
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    ww = int(sw * 0.92)
    wh = int(sh * 0.88)
    win.geometry(f"{ww}x{wh}")
  except Exception:
    win.geometry("1180x900")
  win.minsize(980, 720)
  try:
    win.attributes("-topmost", True)
  except Exception:
    pass

  outer = tk.Frame(win, bg=UI_BG)
  outer.pack(fill="both", expand=True)

  header = tk.Frame(outer, bg=UI_NAVY, height=78)
  header.pack(fill="x")
  header.pack_propagate(False)
  tk.Label(header, text="MOORING FEASIBILITY — FAILURE CRITERIA",
           bg=UI_NAVY, fg="white", font=("Segoe UI", 16, "bold")).pack(
               anchor="w", padx=20, pady=(13, 1))
  tk.Label(header, text="Live report — updates automatically with the selected configuration",
           bg=UI_NAVY, fg="#dce6ed", font=("Segoe UI", 9)).pack(anchor="w", padx=21)

  meta = tk.Frame(outer, bg="white", height=46, bd=1, relief="solid")
  meta.pack(fill="x", padx=14, pady=(12, 8))
  meta.pack_propagate(False)
  config_lbl = tk.Label(meta, text="Configuration: --", bg="white", fg=UI_TEXT,
                        font=("Segoe UI", 10, "bold"))
  config_lbl.pack(side="left", padx=14, pady=10)
  status_lbl = tk.Label(meta, text="STATUS: --", bg="white", fg=UI_TEXT,
                        font=("Segoe UI", 10, "bold"))
  status_lbl.pack(side="right", padx=14)

  canvas = tk.Canvas(outer, bg=UI_BG, highlightthickness=0)
  scrollbar = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
  canvas.configure(yscrollcommand=scrollbar.set)
  scrollbar.pack(side="right", fill="y")
  canvas.pack(side="left", fill="both", expand=True, padx=(12, 0), pady=(0, 12))
  content = tk.Frame(canvas, bg=UI_BG)
  canvas_window = canvas.create_window((0, 0), window=content, anchor="nw")

  def _on_content_config(event=None):
    try:
      visible_w = max(int(canvas.winfo_width() - 4), 1)
      canvas.itemconfigure(canvas_window, width=visible_w)
    except Exception:
      pass
    try:
      canvas.configure(scrollregion=canvas.bbox("all"))
    except Exception:
      pass
  content.bind("<Configure>", _on_content_config)
  canvas.bind("<Configure>", _on_content_config)

  def _on_mousewheel(event):
    try:
      delta = getattr(event, "delta", 0)
      if delta:
        step = -int(delta / 120)
        if step == 0:
          step = -1 if delta > 0 else 1
        canvas.yview_scroll(step, "units")
      elif getattr(event, "num", None) == 4:
        canvas.yview_scroll(-3, "units")
      elif getattr(event, "num", None) == 5:
        canvas.yview_scroll(3, "units")
    except Exception:
      pass

  win.bind("<MouseWheel>", _on_mousewheel, add="+")
  win.bind("<Button-4>", _on_mousewheel, add="+")
  win.bind("<Button-5>", _on_mousewheel, add="+")
  canvas.bind("<MouseWheel>", _on_mousewheel, add="+")
  canvas.bind("<Button-4>", _on_mousewheel, add="+")
  canvas.bind("<Button-5>", _on_mousewheel, add="+")

  panels = {}

  def _panel(name, title, row, col, colspan=1, height=185):
    frame = tk.Frame(content, bg="white", bd=1, relief="solid")
    frame.grid(row=row, column=col, columnspan=colspan, sticky="nsew", padx=7, pady=7)
    head = tk.Frame(frame, bg=UI_PANEL, height=38)
    head.pack(fill="x")
    head.pack_propagate(False)
    head_label = tk.Label(head, text=title, bg=UI_PANEL, fg=UI_TEXT,
             font=("Segoe UI", 10, "bold"))
    head_label.pack(anchor="w", padx=12, pady=9)
    body = tk.Canvas(frame, height=height, bg="white", highlightthickness=0)
    body.pack(fill="both", expand=True, padx=4, pady=4)
    panels[name] = (frame, body, head_label)
    return body

  for c in range(2):
    content.grid_columnconfigure(c, weight=1, uniform="report_col")

  c_geom = _panel("geometry", "01  |  LINE GEOMETRY", 0, 0, height=215)
  c_ground = _panel("ground", "02  |  GROUNDED / SUSPENDED LINE", 0, 1, height=215)
  c_angle = _panel("angle", "03  |  SUSPENDED LINE ANGLE", 1, 0, height=215)
  c_tension = _panel("tension", "04  |  MAXIMUM TENSION LIMITS", 1, 1, height=285)
  c_break = _panel("break", "05  |  BREAK LOAD / MBL", 2, 0, height=325)
  c_fos = _panel("fos", "06  |  ANCHOR & LINE FoS", 2, 1, height=255)
  c_status = _panel("status", "07  |  CURRENT FAILURE STATUS", 3, 0, colspan=2, height=230)
  c_raster = _panel("raster", "08  |  RASTER DEPTH SUITABILITY", 4, 0, colspan=2, height=250)
  c_tension_plots = _panel("tension_plots", "09  |  ANCHOR & FAIRLEAD TENSION COMPONENTS", 5, 0, colspan=2, height=360)

  def _clear(c):
    c.delete("all")

  def _text(c, x, y, text, size=9, weight="normal", fill="#263238", anchor="w"):
    c.create_text(x, y, text=text, font=("Segoe UI", size, weight), fill=fill, anchor=anchor)

  def _bar(c, x0, x1, y, value, low=None, high=None, vmax=100.0, label="", suffix="", marker_color=None):
    value = float(np.clip(value, 0.0, vmax)) if np.isfinite(value) else 0.0
    if marker_color is None:
      marker_color = UI_ACCENT
    c.create_rectangle(x0, y-8, x1, y+8, fill="#e5e9ed", outline="#c8d0d6")
    xpos = x0 + (x1-x0) * value / vmax
    c.create_rectangle(x0, y-8, xpos, y+8, fill=marker_color, outline="")
    if low is not None:
      lx = x0 + (x1-x0) * low / vmax
      c.create_line(lx, y-14, lx, y+14, fill=UI_WARN, width=2)
    if high is not None:
      hx = x0 + (x1-x0) * high / vmax
      c.create_line(hx, y-14, hx, y+14, fill=UI_FAIL, width=2)
    _text(c, x0, y+25, "0", 7, fill="#77858f")
    _text(c, x1, y+25, f"{vmax:g}{suffix}", 7, fill="#77858f", anchor="e")
    if label:
      _text(c, (x0+x1)/2, y-25, label, 8, "bold", anchor="center")
    return xpos

  def _criterion_colour(ok):
    return UI_PASS if bool(ok) else UI_FAIL

  def _numeric(value):
    try:
      v = float(value)
      return v if np.isfinite(v) else np.nan
    except Exception:
      return np.nan

  def _derive_geometry_values(data, system):
    actual = _numeric(data.get("line_length", data.get("line_length_main", np.nan)))
    minimum = _numeric(data.get("min_geometric_length", np.nan))
    if str(system).lower() == "taut":
      # The Taut solver deliberately generates its optimal/stretched line length
      # from the direct padeye-to-fairlead span.  Report that generated value as
      # the Taut design minimum so the criterion compares like-with-like.  The
      # true direct geometric span is shown separately as diagnostic information.
      if np.isfinite(actual):
        minimum = actual
      elif not np.isfinite(minimum):
        xf = _numeric(data.get("xf", np.nan))
        z_anchor = _numeric(data.get("z_anchor_custom", data.get("padeye_z", np.nan)))
        fairlead_draft = _numeric(data.get("fairlead_draft", np.nan))
        if np.isfinite(xf) and np.isfinite(z_anchor) and np.isfinite(fairlead_draft):
          minimum = float(np.hypot(xf, fairlead_draft - z_anchor))
    excess = _numeric(data.get("excess_line", np.nan))
    if str(system).lower() == "taut" and np.isfinite(actual) and np.isfinite(minimum):
      excess = 0.0
    elif not np.isfinite(excess) and np.isfinite(actual) and np.isfinite(minimum):
      excess = actual - minimum
    return actual, minimum, excess

  def _derive_ground_values(data):
    gr = _numeric(data.get("ground_ratio", np.nan))
    lg = _numeric(data.get("grounded_surface_length", data.get("L_bot", np.nan)))
    ls = _numeric(data.get("suspended_surface_length", data.get("L_sus", np.nan)))
    if not np.isfinite(gr) and np.isfinite(lg) and np.isfinite(ls) and (lg + ls) > 0:
      gr = lg / (lg + ls)
    return gr * 100.0 if np.isfinite(gr) else np.nan, lg, ls

  def _derive_angle(data, system):
    angle = _numeric(data.get("suspended_line_angle_deg", np.nan))
    if np.isfinite(angle):
      return angle
    xf = _numeric(data.get("xf", np.nan))
    fairlead_draft = _numeric(data.get("fairlead_draft", np.nan))
    if str(system).lower() == "taut":
      z_anchor = _numeric(data.get("z_anchor_custom", data.get("padeye_z", np.nan)))
      if np.isfinite(xf) and np.isfinite(z_anchor) and np.isfinite(fairlead_draft):
        return float(np.degrees(np.arctan2(abs(fairlead_draft - z_anchor), max(abs(xf), 1e-12))))
      return np.nan
    x_td = _numeric(data.get("X_td", np.nan))
    water_depth = abs(_numeric(data.get("water_depth", np.nan)))
    if np.isfinite(xf) and np.isfinite(x_td) and np.isfinite(fairlead_draft) and np.isfinite(water_depth):
      run = abs(xf - x_td)
      rise = abs(fairlead_draft + water_depth)
      return float(np.degrees(np.arctan2(rise, max(run, 1e-12))))
    return np.nan

  def _segmented_ratio(c, x0, x1, y, ground_pct):
    ground_pct = float(np.clip(ground_pct, 0.0, 100.0))
    split = x0 + (x1-x0)*ground_pct/100.0
    c.create_rectangle(x0, y-14, split, y+14, fill="#748c99", outline="")
    c.create_rectangle(split, y-14, x1, y+14, fill="#b9d7dd", outline="")
    for pct, col in ((15.0, UI_WARN), (85.0, UI_FAIL)):
      xx = x0 + (x1-x0)*pct/100.0
      c.create_line(xx, y-20, xx, y+20, fill=col, width=2)
      _text(c, xx, y+31, f"{pct:.0f}%", 7, "bold", col, anchor="center")
    _text(c, x0, y-29, "GROUNDED", 8, "bold", "#536873")
    _text(c, x1, y-29, "SUSPENDED", 8, "bold", "#536873", anchor="e")
    return split

  def _panel_meta(name):
    return panels[name]

  def _show_failure_fallback(c, reason, y=76, width=467, box_height=58, reason_below=True, scope=None):
    # Full fallback box retained for the status panel only. Criterion panels use
    # _show_bar_not_evaluated so the placeholder replaces the missing visual.
    reason = str(reason or "The configuration failed before this criterion could be calculated.")
    x0 = max(14, (500 - width) / 2.0)
    x1 = min(500 - 14, x0 + width)
    actual_width = x1 - x0
    c.create_rectangle(x0, y, x1, y + box_height, fill="#f3f5f7", outline="#c9d1d6")
    title = "NOT EVALUATED" if not scope else f"NOT EVALUATED — {scope}"
    c.create_text((x0+x1)/2, y+13, text=title, width=max(actual_width-16, 40),
                  font=("Segoe UI", 8, "bold"), fill="#7a8791", anchor="center")
    c.create_text((x0+x1)/2, y+33,
                  text="Upstream solver failure — this criterion was not calculated.",
                  width=max(actual_width-16, 40), font=("Segoe UI", 7), fill="#667781", anchor="center")
    if reason_below:
      c.create_text((x0+x1)/2, y + box_height - 6, text=reason, width=max(actual_width-16, 40),
                    font=("Segoe UI", 6), fill=UI_FAIL, anchor="s")

  def _show_bar_not_evaluated(c, y, x0=28, x1=485, width=250, height=46, scope=None):
    """Replace the missing visual bar with a compact, clearly labelled placeholder."""
    width = min(float(width), float(x1-x0))
    mid = (x0+x1)/2.0
    box_x0 = mid - width/2.0
    box_x1 = mid + width/2.0
    box_y0 = y - height/2.0
    box_y1 = y + height/2.0
    c.create_rectangle(box_x0, box_y0, box_x1, box_y1,
                        fill="#f3f5f7", outline="#c9d1d6")
    title = "NOT EVALUATED" if not scope else f"NOT EVALUATED — {scope}"
    c.create_text(mid, y-8, text=title, width=max(width-12, 40),
                  font=("Segoe UI", 7, "bold"), fill="#7a8791", anchor="center")
    c.create_text(mid, y+11, text="Upstream solver failure — no valid value was calculated.",
                  width=max(width-12, 40), font=("Segoe UI", 6), fill="#667781", anchor="center")

  def _failure_reason_text(data):
    reasons = []
    msg = str(data.get("msg", "") or "").strip()
    if msg and msg.lower() not in ("none", "--"):
      reasons.append(msg)
    for r in (data.get("_raster_failure_reasons") or [])[:4]:
      rs = str(r).strip()
      if rs and rs not in reasons:
        reasons.append(rs)
    return "\n".join(reasons) if reasons else "Configuration failed before this criterion generated a numerical result."

  def _apply_panel_layout(system, data):
    # Hide criteria that do not apply to the selected mooring type and
    # compact the remaining panels into a two-column report.
    visible = ["geometry"]
    if str(system).lower() != "taut":
      visible.extend(["ground", "angle"])
    # Ground-line ratio and suspended-line angle are not failure criteria for a fully taut system.
    visible.extend(["tension", "break", "fos", "status", "raster", "tension_plots"])

    for name, (frame, _body, _head) in panels.items():
      frame.grid_remove()

    row = 0
    col = 0
    full_width_panels = {"status", "raster", "tension_plots"}
    for idx, name in enumerate(visible):
      frame, _body, head = panels[name]
      if name in full_width_panels:
        # Start every spanning panel on a fresh row.  This avoids a wide
        # status/raster panel sharing a grid row with the preceding half-width
        # criterion panel when the visible set changes by mooring type.
        if col:
          row += 1
          col = 0
        frame.grid(row=row, column=0, columnspan=2, sticky="nsew", padx=7, pady=7)
        row += 1
      else:
        frame.grid(row=row, column=col, sticky="nsew", padx=7, pady=7)
        col += 1
        if col == 2:
          row += 1
          col = 0
      head.configure(text=f"{idx+1:02d}  |  " + {
        "geometry":"LINE GEOMETRY",
        "ground":"GROUNDED / SUSPENDED LINE",
        "angle":"SUSPENDED LINE ANGLE",
        "tension":"MAXIMUM TENSION LIMITS",
        "break":"BREAK LOAD / MBL",
        "fos":"ANCHOR & LINE FoS",
        "status":"CURRENT FAILURE STATUS",
        "raster":"RASTER DEPTH SUITABILITY",
        "tension_plots":"ANCHOR & FAIRLEAD TENSION COMPONENTS",
      }[name])
    if col:
      row += 1
    for r in range(0, row + 1):
      content.grid_rowconfigure(r, weight=0)
    return set(visible)


  def _draw_tension_component_plot(c, x0, y0, w, h, H, V, allowed, title, anchor_type=""):
    """Draw an H/V tension envelope. DEA anchor plots enforce V=0 by definition."""
    c.create_text(x0 + w/2, y0 + 8, text=title, font=("Segoe UI", 9, "bold"),
                  fill=UI_TEXT, anchor="n")
    left, top = x0 + 55, y0 + 38
    right, bottom = x0 + w - 18, y0 + h - 44
    H = _numeric(H); V = _numeric(V); allowed = _numeric(allowed)
    is_dea_anchor = title.upper().startswith("ANCHOR") and str(anchor_type).strip().lower() == "dea"

    # DEA tension is NOT forced to zero for reporting.  The actual signed
    # vertical component must be shown because positive V represents uplift
    # and is a failure condition for a DEA.
    current = float(np.hypot(H, V)) if np.isfinite(H) and np.isfinite(V) else np.nan
    radius = allowed if np.isfinite(allowed) and allowed > 0 else (current if np.isfinite(current) else 1.0)
    vmax = max(radius, abs(H) if np.isfinite(H) else 0.0, abs(V) if np.isfinite(V) else 0.0, 1.0) * 1.15
    plot_w = right - left
    plot_h = bottom - top
    sx = plot_w / vmax
    # Symmetric vertical scale so that positive/negative V are both meaningful.
    sy = (plot_h / 2.0) / vmax
    v_mid = (top + bottom) / 2.0

    c.create_rectangle(left, top, right, bottom, fill="#fff0f0", outline="#d2d9df")

    if is_dea_anchor:
      # DEA admissible region: no upward/uplift vertical reaction.
      # Negative/zero V is acceptable; positive V is the failure side.
      split_y = v_mid
      c.create_rectangle(left, split_y, right, bottom, fill="#e7f4ea", outline="")
      c.create_rectangle(left, top, right, split_y, fill="#fff0f0", outline="")
      c.create_line(left, split_y, right, split_y, fill=UI_PASS, width=1.8)
      c.create_text(left+8, top+5,
                    text="UPLIFT / FAIL: V > 0",
                    font=("Segoe UI", 7, "bold"), fill=UI_FAIL, anchor="nw")
      c.create_text(left+8, bottom-5,
                    text="ACCEPTABLE: V ≤ 0",
                    font=("Segoe UI", 7, "bold"), fill=UI_PASS, anchor="sw")
    elif np.isfinite(allowed) and allowed > 0:
      # Full resultant-tension allowable circle, clipped to the H/V plotting
      # rectangle. This gives a clear green allowable zone and red exceedance.
      xs = np.linspace(-allowed, allowed, 240)
      ys = np.sqrt(np.maximum(allowed**2 - xs**2, 0.0))
      upper = [(left + (x+vmax)*sx/2.0, v_mid - y*sy) for x, y in zip(xs, ys)]
      lower = [(left + (x+vmax)*sx/2.0, v_mid + y*sy) for x, y in zip(xs[::-1], ys[::-1])]
      poly = upper + lower
      c.create_polygon(poly, fill="#e7f4ea", outline="")

    # Axes: H horizontal, V vertical, with V=0 crossing the middle.
    c.create_line(left, v_mid, right, v_mid, fill="#596a74", width=1.2)
    c.create_line(left, bottom, left, top, fill="#596a74", width=1.2)

    for frac in (-1.0, -0.5, 0.0, 0.5, 1.0):
      xx = left + plot_w*(frac + 1.0)/2.0
      c.create_line(xx, v_mid-5, xx, v_mid+5, fill="#7d8b93")
      if frac != 0.0:
        c.create_text(xx, v_mid+8, text=f"{frac*vmax/1000.0:.0f}",
                      font=("Segoe UI", 6), fill="#667781", anchor="n")
    for frac in (-1.0, -0.5, 0.0, 0.5, 1.0):
      yy = v_mid - frac*(plot_h/2.0)
      c.create_line(left, yy, left+5, yy, fill="#7d8b93")
      c.create_text(left-7, yy, text=f"{frac*vmax/1000.0:.0f}",
                    font=("Segoe UI", 6), fill="#667781", anchor="e")

    c.create_text((left+right)/2, bottom+22, text="Horizontal tension, H (kN)",
                  font=("Segoe UI", 7), fill="#52626c", anchor="n")
    c.create_text(left-38, v_mid, text="Vertical tension, V (kN)",
                  font=("Segoe UI", 7), fill="#52626c", angle=90, anchor="center")

    if is_dea_anchor:
      c.create_text(right-6, v_mid+5,
                    text="DEA limit: V ≤ 0 kN",
                    font=("Segoe UI", 7), fill=UI_PASS, anchor="ne")
    elif np.isfinite(allowed) and allowed > 0:
      c.create_text(left+6, top+4, text=f"Allowable resultant: {allowed/1000:.1f} kN",
                    font=("Segoe UI", 7), fill=UI_PASS, anchor="sw")

    if np.isfinite(H) and np.isfinite(V):
      # H is plotted from zero-centered negative-to-positive axis. Anchor/fairlead
      # components are normally positive, but the visual remains general.
      px = left + (float(np.clip(H, -vmax, vmax)) + vmax) * sx / 2.0
      py = v_mid - float(np.clip(V, -vmax, vmax)) * sy
      if is_dea_anchor:
        # Positive V is uplift and therefore fails the DEA criterion.
        point_ok = V <= 1e-6
      else:
        point_ok = (not np.isfinite(allowed)) or current <= allowed + 1e-9
      pc = UI_PASS if point_ok else UI_FAIL
      c.create_oval(px-5, py-5, px+5, py+5, fill=pc, outline="white", width=1)
      label_x = min(max(px+8, left+8), right-5)
      label_y = max(min(py-8, bottom-8), top+8)
      v_label = V
      c.create_text(label_x, label_y,
                    text=f"H={H/1000:.1f} kN\nV={v_label/1000:.1f} kN",
                    font=("Segoe UI", 7), fill=pc, anchor="sw")

    c.create_text(left+4, bottom+7, text="negative", font=("Segoe UI", 6), fill="#8a949b", anchor="nw")
    c.create_text(right-4, bottom+7, text="positive", font=("Segoe UI", 6), fill="#8a949b", anchor="ne")

  last_report_config = {"key": None}

  def refresh():
    try:
      payload = data_provider() if data_provider is not None else {}
      payload = payload or {}
      data = payload.get("data") or {}
      system = payload.get("system", data.get("system_type", "--"))
      anchor = payload.get("anchor", data.get("primary_anc", "--"))
      profile = payload.get("profile", "")
      config_key = (str(system), str(anchor), str(profile))
      config_changed = last_report_config["key"] is not None and last_report_config["key"] != config_key
      first_render = last_report_config["key"] is None
      last_report_config["key"] = config_key
      config_lbl.configure(text=f"Configuration: {system}  |  {anchor}" + (f"  |  {profile}" if profile else ""))

      visible_panels = _apply_panel_layout(system, data)
      failure_reason = _failure_reason_text(data)
      evals = failure_evaluation_flags(system, data)
      failure_stage = evals["stage"]

      base_ok = bool(data.get("success"))
      raster_ok = data.get("_raster_system_pass", True)
      overall_ok = base_ok and bool(raster_ok)
      status_lbl.configure(text="STATUS: PASS" if overall_ok else "STATUS: FAIL",
                           fg=UI_PASS if overall_ok else UI_FAIL)

      # 01 Geometry
      c = c_geom; _clear(c)
      actual, minimum, excess = _derive_geometry_values(data, system)
      _text(c, 18, 20, f"Actual line: {actual:.2f} m" if np.isfinite(actual) else "Actual line: --", 9, "bold")
      if str(system).lower() == "taut":
        direct_span_diag = _numeric(data.get("direct_geometric_span", np.nan))
        _text(c, 18, 43, f"Taut design minimum: {minimum:.2f} m" if np.isfinite(minimum) else "Taut design minimum: --")
        if np.isfinite(direct_span_diag):
          _text(c, 18, 66, f"Direct geometric span: {direct_span_diag:.2f} m")
        else:
          _text(c, 18, 66, f"Excess line: {excess:+.2f} m" if np.isfinite(excess) else "Excess line: --")
      else:
        _text(c, 18, 43, f"Minimum geometric: {minimum:.2f} m" if np.isfinite(minimum) else "Minimum geometric: --")
        _text(c, 18, 66, f"Excess line: {excess:+.2f} m" if np.isfinite(excess) else "Excess line: --")
      if np.isfinite(actual) and np.isfinite(minimum) and minimum > 0:
        vmax = max(actual, minimum, 1.0)*1.15
        geom_ok = actual + 1e-6 >= minimum
        _bar(c, 28, 485, 105, actual, high=minimum, vmax=vmax,
             label="Actual line length vs geometric minimum", suffix=" m",
             marker_color=_criterion_colour(geom_ok))
        _text(c, 28, 150, "Taut design length is the generated optimal length; direct padeye-to-fairlead span is shown above.",
              8, fill="#667781")
        _text(c, 28, 174, "PASS: generated taut length is the design minimum." if geom_ok else "FAIL: no valid generated taut length.",
              8, fill=UI_PASS if geom_ok else UI_FAIL)
      else:
        _show_bar_not_evaluated(c, 105, scope="LINE GEOMETRY")
      if np.isfinite(actual) and np.isfinite(minimum) and actual + 1e-6 < minimum:
        _text(c, 28, 198, "UPSTREAM FAILURE: downstream tension / FoS checks are only reported when their inputs were successfully calculated.",
              8, "bold", UI_FAIL)

      # 02 Grounded / suspended
      c = c_ground
      if "ground" in visible_panels:
        _clear(c)
        gr, lg, ls = _derive_ground_values(data) if evals["ground"] else (np.nan, np.nan, np.nan)
        _text(c, 18, 20, f"Grounded: {lg:.2f} m" if np.isfinite(lg) else "Grounded: --", 9, "bold")
        _text(c, 185, 20, f"Suspended: {ls:.2f} m" if np.isfinite(ls) else "Suspended: --", 9, "bold")
        _text(c, 18, 43, f"Ground ratio: {gr:.1f}%" if np.isfinite(gr) else "Ground ratio: --", 9, "bold")
        if np.isfinite(gr):
          # The bar is green when the combined ground-line criterion passes and
          # red when either the <15% or the >85%+near-vertical condition fails.
          angle_for_ground = _derive_angle(data, system)
          ground_low_fail = gr < (GROUND_LINE_MIN_RATIO * 100.0)
          ground_high_fail = (
              gr > (GROUND_LINE_MAX_RATIO * 100.0)
              and np.isfinite(angle_for_ground)
              and angle_for_ground >= GROUND_LINE_VERTICAL_ANGLE_DEG
          )
          ground_ok = not (ground_low_fail or ground_high_fail)
          _text(c, 405, 20, "PASS" if ground_ok else "FAIL", 8, "bold", UI_PASS if ground_ok else UI_FAIL, anchor="e")
          # Keep the grounded/suspended split visible, but colour the active
          # grounded portion by the criterion state so its status is obvious.
          split = float(np.clip(gr, 0.0, 100.0))
          _segmented_ratio(c, 28, 485, 100, split)
          # Overlay the current grounded portion with pass/fail colour.
          px = 28 + (485-28)*split/100.0
          if px > 28:
            c.create_rectangle(28, 86, px, 114, fill=_criterion_colour(ground_ok), outline="")
          _text(c, 28, 160, "<15% = FAIL", 8, fill=UI_WARN)
          _text(c, 28, 182, ">85% = FAIL only when suspended line is near-vertical", 8, fill=UI_FAIL)
        else:
          _show_bar_not_evaluated(c, 100, scope="GROUND / SUSPENDED")

      # 03 Angle
      c = c_angle; _clear(c)
      angle = _derive_angle(data, system) if evals["angle"] else np.nan
      _text(c, 18, 22, f"Current suspended angle: {angle:.2f}° from horizontal" if np.isfinite(angle) else "Suspended angle: --", 9, "bold")
      _text(c, 18, 47, f"Near-vertical threshold: {GROUND_LINE_VERTICAL_ANGLE_DEG:.1f}°", 9)
      if np.isfinite(angle):
        gr_angle, _, _ = _derive_ground_values(data)
        angle_fail = (
            str(system).lower() != "taut"
            and np.isfinite(gr_angle)
            and gr_angle > (GROUND_LINE_MAX_RATIO * 100.0)
            and angle >= GROUND_LINE_VERTICAL_ANGLE_DEG
        )
        _text(c, 405, 22, "FAIL" if angle_fail else "PASS", 8, "bold", UI_FAIL if angle_fail else UI_PASS, anchor="e")
        _bar(c, 28, 485, 108, angle, high=GROUND_LINE_VERTICAL_ANGLE_DEG, vmax=90.0,
             label="0° horizontal  →  90° vertical", suffix="°",
             marker_color=_criterion_colour(not angle_fail))
        _text(c, 28, 157, "Used with the >85% grounded criterion.", 8, fill="#667781")
        _text(c, 28, 180, "PASS" if not angle_fail else "FAIL: near-vertical angle combined with excessive ground line.",
              8, fill=UI_PASS if not angle_fail else UI_FAIL)
      else:
        _show_bar_not_evaluated(c, 108, scope="SUSPENDED ANGLE")

      # 04 Maximum tension limits — MBL / required FoS
      c = c_tension; _clear(c)
      max_t = _numeric(data.get("T_max", np.nan))
      rope_tmax = _numeric(data.get("taut_max_tension", np.nan))
      chain_d = _numeric(data.get("chain_d", np.nan))
      rope_d = _numeric(data.get("rope_d", np.nan))
      chain_mbl = np.nan
      rope_mbl = np.nan
      chain_allowed = np.nan
      rope_allowed = np.nan
      if np.isfinite(chain_d):
        grade = str(data.get("chain_grade", "R4"))
        chain_mbl = calculate_chain_mbl(chain_d, grade)
        chain_allowed = chain_mbl / max(MIN_LINE_MBL_FOS, 1e-9)
      if np.isfinite(rope_d):
        material = str(data.get("rope_material", "Polyester" if system == "Semi-Taut" else "HMPE"))
        rope_mbl = calculate_rope_mbl(rope_d, material)
        rope_allowed = rope_mbl / max(MIN_LINE_MBL_FOS, 1e-9)

      drew_tension = False
      tension_states = []
      tension_inputs_available = False
      if not evals["tension"]:
        # Static allowable values may still be displayed, but the comparison bar
        # is explicitly not evaluated because an upstream stage failed first.
        if system == "Catenary" and np.isfinite(chain_allowed):
          _text(c, 18, 20, f"Maximum chain tension allowed: {chain_allowed/1000.0:.2f} kN", 9, "bold")
          _text(c, 220, 20, "Current governing: --", 9)
          _show_bar_not_evaluated(c, 95, scope="CHAIN TENSION")
        elif system == "Semi-Taut":
          y = 20
          if np.isfinite(chain_allowed):
            _text(c, 18, y, f"Chain max allowed: {chain_allowed/1000.0:.2f} kN", 9, "bold")
            _text(c, 220, y, "Current governing: --", 9)
            _show_bar_not_evaluated(c, 72, scope="CHAIN TENSION")
            y = 122
          if np.isfinite(rope_allowed):
            _text(c, 18, y, f"Rope max allowed: {rope_allowed/1000.0:.2f} kN", 9, "bold")
            _text(c, 220, y, "Current rope max: --", 9)
            _show_bar_not_evaluated(c, y+52, scope="ROPE TENSION")
        elif system == "Taut" and np.isfinite(rope_allowed):
          _text(c, 18, 20, f"Maximum rope tension allowed: {rope_allowed/1000.0:.2f} kN", 9, "bold")
          _text(c, 18, 43, "Current maximum tension: --", 9)
          _show_bar_not_evaluated(c, 95, scope="ROPE TENSION")
        else:
          _show_bar_not_evaluated(c, 95, scope="TENSION LIMITS")
      elif system == "Catenary" and np.isfinite(chain_allowed):
        chain_cur = max_t
        _text(c, 18, 20, f"Maximum chain tension allowed: {chain_allowed/1000.0:.2f} kN", 9, "bold")
        _text(c, 18, 43, f"Current governing tension: {chain_cur/1000.0:.2f} kN" if np.isfinite(chain_cur) else "Current governing tension: --", 9)
        if np.isfinite(chain_cur):
          tension_inputs_available = True
          cur = chain_cur/1000.0
          high = chain_allowed/1000.0
          vmax = max(high, cur, 1.0)*1.20
          state_ok = chain_cur <= chain_allowed
          _bar(c, 28, 485, 95, cur, high=high, vmax=vmax,
               label="Current tension vs maximum allowed", suffix=" kN",
               marker_color=_criterion_colour(state_ok))
          tension_states.append(bool(state_ok))
          _text(c, 405, 20, "PASS" if state_ok else "FAIL", 8, "bold", UI_PASS if state_ok else UI_FAIL, anchor="e")
          _text(c, 18, 148, "Maximum allowed = chain MBL / required line FoS.", 8, fill="#667781")
          drew_tension = True
      elif system == "Semi-Taut":
        y = 20
        if np.isfinite(chain_allowed):
          chain_cur = max_t
          _text(c, 18, y, f"Chain max allowed: {chain_allowed/1000.0:.2f} kN", 9, "bold")
          _text(c, 220, y, f"Current governing: {chain_cur/1000.0:.2f} kN" if np.isfinite(chain_cur) else "Current governing: --", 9)
          if np.isfinite(chain_cur):
            tension_inputs_available = True
            cur = chain_cur/1000.0
            high = chain_allowed/1000.0
            vmax = max(high, cur, 1.0)*1.20
            state_ok = chain_cur <= chain_allowed
            _bar(c, 28, 485, 72, cur, high=high, vmax=vmax,
                 label="Chain current vs maximum allowed", suffix=" kN",
                 marker_color=_criterion_colour(state_ok))
            tension_states.append(bool(state_ok))
            drew_tension = True
          y = 122
        if np.isfinite(rope_allowed):
          rope_cur = rope_tmax if np.isfinite(rope_tmax) else max_t
          _text(c, 18, y, f"Rope max allowed: {rope_allowed/1000.0:.2f} kN", 9, "bold")
          _text(c, 220, y, f"Current rope max: {rope_cur/1000.0:.2f} kN" if np.isfinite(rope_cur) else "Current rope max: --", 9)
          if np.isfinite(rope_cur):
            tension_inputs_available = True
            cur = rope_cur/1000.0
            high = rope_allowed/1000.0
            vmax = max(high, cur, 1.0)*1.20
            state_ok = rope_cur <= rope_allowed
            _bar(c, 28, 485, y+52, cur, high=high, vmax=vmax,
                 label="Rope current vs maximum allowed", suffix=" kN",
                 marker_color=_criterion_colour(state_ok))
            tension_states.append(bool(state_ok))
            drew_tension = True
      elif system == "Taut" and np.isfinite(rope_allowed):
        rope_cur = rope_tmax if np.isfinite(rope_tmax) else max_t
        _text(c, 18, 20, f"Maximum rope tension allowed: {rope_allowed/1000.0:.2f} kN", 9, "bold")
        _text(c, 18, 43, f"Current maximum tension: {rope_cur/1000.0:.2f} kN" if np.isfinite(rope_cur) else "Current maximum tension: --", 9)
        if np.isfinite(rope_cur):
          tension_inputs_available = True
          cur = rope_cur/1000.0
          high = rope_allowed/1000.0
          vmax = max(high, cur, 1.0)*1.20
          state_ok = rope_cur <= rope_allowed
          _bar(c, 28, 485, 95, cur, high=high, vmax=vmax,
               label="Current tension vs maximum allowed", suffix=" kN",
               marker_color=_criterion_colour(state_ok))
          tension_states.append(bool(state_ok))
          _text(c, 405, 20, "PASS" if state_ok else "FAIL", 8, "bold", UI_PASS if state_ok else UI_FAIL, anchor="e")
          _text(c, 18, 148, "Maximum allowed = rope MBL / required line FoS.", 8, fill="#667781")
          drew_tension = True
      if evals["tension"] and not tension_inputs_available:
        if system == "Semi-Taut":
          if np.isfinite(chain_allowed):
            _show_bar_not_evaluated(c, 72, scope="CHAIN TENSION")
          if np.isfinite(rope_allowed):
            _show_bar_not_evaluated(c, 174, scope="ROPE TENSION")
        elif system == "Catenary" and np.isfinite(chain_allowed):
          _show_bar_not_evaluated(c, 95, scope="CHAIN TENSION")
        elif system == "Taut" and np.isfinite(rope_allowed):
          _show_bar_not_evaluated(c, 95, scope="ROPE TENSION")
        else:
          _show_bar_not_evaluated(c, 95, scope="TENSION LIMITS")
      elif evals["tension"] and not drew_tension:
        _show_bar_not_evaluated(c, 76, scope="TENSION LIMITS")

      # 05 Break load / MBL
      c = c_break; _clear(c)
      max_t = _numeric(data.get("T_max", np.nan))
      if not np.isfinite(max_t) and str(system).lower() == "taut":
        max_t = _numeric(data.get("taut_max_tension", np.nan))
      y = 20
      any_output = False
      if system != "Taut" and np.isfinite(chain_d):
        grade = str(data.get("chain_grade", "R4"))
        chain_mbl = calculate_chain_mbl(chain_d, grade)
        _text(c, 18, y, f"Chain MBL ({grade}): {chain_mbl/1000.0:.2f} kN", 9, "bold")
        req_chain = max_t * MIN_LINE_MBL_FOS if evals["break"] and np.isfinite(max_t) else np.nan
        _text(c, 18, y+23, f"Required minimum BL: {req_chain/1000.0:.2f} kN" if np.isfinite(req_chain) else "Required minimum BL: --", 9)
        if np.isfinite(req_chain):
          vmax = max(chain_mbl, req_chain, 1.0)*1.20
          ok = chain_mbl >= req_chain
          _bar(c, 28, 485, y+68, chain_mbl/1000.0, high=req_chain/1000.0, vmax=vmax/1000.0,
               label="Selected chain MBL vs required break load", suffix=" kN", marker_color=_criterion_colour(ok))
          _text(c, 18, y+97, "PASS" if ok else "FAIL: selected chain MBL is below required break load.", 8, fill=UI_PASS if ok else UI_FAIL)
        else:
          _show_bar_not_evaluated(c, y+68, scope="CHAIN BREAK LOAD")
        any_output = True
        y += 120

      if np.isfinite(rope_d):
        material = str(data.get("rope_material", "Polyester" if system == "Semi-Taut" else "HMPE"))
        rope_mbl = calculate_rope_mbl(rope_d, material)
        _text(c, 18, y, f"{'Taut ' if system == 'Taut' else ''}Rope MBL ({material}): {rope_mbl/1000.0:.2f} kN", 9, "bold")
        req_rope = max_t * MIN_LINE_MBL_FOS if evals["break"] and np.isfinite(max_t) else np.nan
        _text(c, 18, y+23, f"Required minimum BL: {req_rope/1000.0:.2f} kN" if np.isfinite(req_rope) else "Required minimum BL: --", 9)
        if np.isfinite(req_rope):
          vmax = max(rope_mbl, req_rope, 1.0)*1.20
          ok = rope_mbl >= req_rope
          _bar(c, 28, 485, y+68, rope_mbl/1000.0, high=req_rope/1000.0, vmax=vmax/1000.0,
               label="Selected rope MBL vs required break load", suffix=" kN", marker_color=_criterion_colour(ok))
          _text(c, 18, y+97, "PASS" if ok else "FAIL: selected rope MBL is below required break load.", 8, fill=UI_PASS if ok else UI_FAIL)
        else:
          _show_bar_not_evaluated(c, y+68, scope="ROPE BREAK LOAD")
        any_output = True
        y += 120

      if not any_output:
        _show_bar_not_evaluated(c, 100, scope="BREAK LOAD / MBL")

      # 06 FoS
      c = c_fos; _clear(c)
      line_fos_vals = []
      for key in ("line_fos_fairlead", "line_fos_padeye", "line_fos_chain", "line_fos_rope", "line_fos"):
        try:
          v = float(data.get(key, np.nan))
          if np.isfinite(v): line_fos_vals.append(v)
        except Exception:
          pass
      governing_line_fos = min(line_fos_vals) if line_fos_vals else np.nan
      anchor_fos = []
      for aname, res in (data.get("anchor_results") or {}).items():
        try:
          anchor_fos.append((aname, float(res.get("fos", np.nan)), str(res.get("status", "--"))))
        except Exception:
          pass
      min_anchor_fos = min([x[1] for x in anchor_fos if np.isfinite(x[1])], default=np.nan)

      _text(c, 18, 20, f"Required line FoS: {MIN_LINE_MBL_FOS:.2f}", 9, "bold")
      _text(c, 455, 20, "PASS" if np.isfinite(governing_line_fos) and governing_line_fos >= MIN_LINE_MBL_FOS else ("FAIL" if np.isfinite(governing_line_fos) else "NOT EVALUATED"),
            8, "bold", UI_PASS if np.isfinite(governing_line_fos) and governing_line_fos >= MIN_LINE_MBL_FOS else (UI_FAIL if np.isfinite(governing_line_fos) else "#7a8791"), anchor="e")
      _text(c, 18, 43, f"Governing current line FoS: {governing_line_fos:.2f}" if evals["line_fos"] and np.isfinite(governing_line_fos) else "Governing current line FoS: --", 9)
      if evals["line_fos"] and np.isfinite(governing_line_fos):
        ok = governing_line_fos >= MIN_LINE_MBL_FOS
        _bar(c, 28, 485, 88, governing_line_fos, high=MIN_LINE_MBL_FOS,
             vmax=max(2.0, governing_line_fos*1.15), label="Line FoS", suffix="", marker_color=_criterion_colour(ok))
      else:
        _show_bar_not_evaluated(c, 88, scope="LINE FoS")

      _text(c, 18, 135, "Anchor suitability", 8, "bold")
      _text(c, 455, 135, "PASS" if evals["anchor_fos"] and np.isfinite(min_anchor_fos) and min_anchor_fos >= 1.50 else ("FAIL" if evals["anchor_fos"] and np.isfinite(min_anchor_fos) else "NOT EVALUATED"),
            8, "bold", UI_PASS if evals["anchor_fos"] and np.isfinite(min_anchor_fos) and min_anchor_fos >= 1.50 else (UI_FAIL if evals["anchor_fos"] and np.isfinite(min_anchor_fos) else "#7a8791"), anchor="e")
      _text(c, 18, 158, f"Governing anchor FoS: {min_anchor_fos:.2f}  |  required: 1.50" if evals["anchor_fos"] and np.isfinite(min_anchor_fos) else "Governing anchor FoS: --", 9, "bold")
      if evals["anchor_fos"] and np.isfinite(min_anchor_fos):
        ok = min_anchor_fos >= 1.50
        _bar(c, 28, 485, 193, min_anchor_fos, high=1.50,
             vmax=max(2.0, min_anchor_fos*1.15), label="Anchor FoS", suffix="", marker_color=_criterion_colour(ok))
      else:
        _show_bar_not_evaluated(c, 193, scope="ANCHOR FoS")

      # 08 Anchor & fairlead tension component plots
      c = c_tension_plots; _clear(c)
      frame_w = 445
      allowed = np.nan
      if np.isfinite(rope_d):
        material = str(data.get("rope_material", "Polyester" if system == "Semi-Taut" else "HMPE"))
        allowed = calculate_rope_mbl(rope_d, material) / max(MIN_LINE_MBL_FOS, 1e-9)
      elif np.isfinite(chain_d):
        grade = str(data.get("chain_grade", "R4"))
        allowed = calculate_chain_mbl(chain_d, grade) / max(MIN_LINE_MBL_FOS, 1e-9)
      ha_v = _numeric(data.get("HA", np.nan)); va_v = _numeric(data.get("VA", np.nan))
      hf_v = _numeric(data.get("HF", np.nan)); vf_v = _numeric(data.get("VF", np.nan))
      anchor_tension_ok = evals["tension_components"]
      fairlead_tension_ok = evals["tension_components"]
      if anchor_tension_ok:
        _draw_tension_component_plot(c, 8, 0, frame_w, 320, ha_v, va_v, allowed, "ANCHOR / PADEYE TENSION", anchor_type=anchor)
      else:
        c.create_text(8 + frame_w/2.0, 10, text="ANCHOR / PADEYE TENSION", font=("Segoe UI", 9, "bold"), fill=UI_TEXT, anchor="n")
        _show_bar_not_evaluated(c, 160, x0=70, x1=390, width=250, height=44, scope="ANCHOR TENSION")
      if fairlead_tension_ok:
        _draw_tension_component_plot(c, 455, 0, frame_w, 320, hf_v, vf_v, allowed, "FAIRLEAD TENSION", anchor_type=anchor)
      else:
        c.create_text(455 + frame_w/2.0, 10, text="FAIRLEAD TENSION", font=("Segoe UI", 9, "bold"), fill=UI_TEXT, anchor="n")
        _show_bar_not_evaluated(c, 160, x0=517, x1=837, width=250, height=44, scope="FAIRLEAD TENSION")

      # 07 Status, spanning both columns
      c = c_status; _clear(c)
      msg = str(data.get("msg", "Configuration currently passes base feasibility checks."))
      if not base_ok:
        state_text, state_color = "BASE CONFIGURATION: FAIL", UI_FAIL
      elif not bool(raster_ok):
        state_text, state_color = "LOCAL RASTER CHECK: FAIL", UI_FAIL
      else:
        state_text, state_color = "CURRENT CONFIGURATION: PASS", UI_PASS
      _text(c, 18, 22, state_text, 10, "bold", state_color)
      _text(c, 18, 52, "Primary result:", 8, "bold")
      c.create_text(18, 70, text=msg, width=920, font=("Segoe UI", 8), fill="#263238", anchor="nw")
      if not base_ok:
        c.create_text(18, 112, text=(
            "Interpretation: this is the governing failure. Any later section marked NOT EVALUATED "
            "was not calculated because its required upstream inputs were unavailable."
        ), width=920, font=("Segoe UI", 8, "bold"), fill=UI_FAIL, anchor="nw")
      else:
        c.create_text(18, 112, text=(
            "Interpretation: later sections are shown where their numerical inputs were successfully calculated; "
            "they are not automatically failed by an earlier criterion."
        ), width=920, font=("Segoe UI", 8), fill="#667781", anchor="nw")
      diag_note = str(data.get("diagnostic_note", "")).strip()
      if diag_note:
        _text(c, 18, 128, diag_note, 8, fill="#667781")
      if data.get("_raster_turbine_results"):
        _text(
            c, 18, 166 if diag_note else 145,
            "Per-turbine raster-depth status is summarised in section 08 below.",
            8, fill="#667781",
        )
      else:
        _text(c, 18, 145, "No local raster-depth result is currently recorded.", 8, fill="#667781")

      # 08 Raster depth suitability.  This replaces the old floating 3-D text
      # box and groups bridle branches under their owning turbine so a shared
      # centre-depth failure is not repeated for every TnLn branch.
      c = c_raster; _clear(c)
      raster_summary = build_raster_depth_summary(data)
      _text(c, 18, 20, raster_summary["band_text"], 8, "bold", UI_TEXT)
      turbine_summaries = raster_summary["turbines"]
      if not turbine_summaries:
        c.configure(height=170)
        _text(c, 18, 58, "No raster-aware turbine results are available for this configuration.",
              8, fill="#667781")
      else:
        y = 54
        for turbine_summary in turbine_summaries:
          passed = bool(turbine_summary["passed"])
          state_colour = UI_PASS if passed else UI_FAIL
          state_word = "PASS" if passed else "FAIL"
          _text(
              c, 18, y,
              f"T{turbine_summary['number']}  |  {state_word}  |  "
              f"line-anchor depths: {turbine_summary['depth_text']}",
              9, "bold", state_colour,
          )
          issues = turbine_summary["issues"]
          if issues:
            detail = "  •  ".join(issues)
            line_count = max(1, min(4, int(np.ceil(len(detail) / 105.0))))
            c.create_text(
                34, y + 18, text=detail, width=880,
                font=("Segoe UI", 8), fill="#4f5f66", anchor="nw",
            )
            y += 24 + (line_count * 15)
          else:
            _text(c, 34, y + 18, "All centre, anchor and bridle-anchor depth checks are within the Excel H30 band.",
                  8, fill="#4f5f66")
            y += 38
          c.create_line(18, y - 7, 920, y - 7, fill="#e1e6ea")
          y += 5
        c.configure(height=int(np.clip(y + 12, 170, 520)))

      # Recompute the embedded content width after the visible panel set changes.
      # When a new configuration is selected, return the report to the top so
      # section 01 is always visible. The remaining applicable sections remain
      # available below via the scrollbar/mouse-wheel.
      win.update_idletasks()
      try:
        canvas.itemconfigure(canvas_window, width=max(int(canvas.winfo_width() - 4), 1))
      except Exception:
        pass
      win.update_idletasks()
      canvas.configure(scrollregion=canvas.bbox("all"))
      if first_render or config_changed:
        canvas.yview_moveto(0.0)
      win.update_idletasks()
      canvas.configure(scrollregion=canvas.bbox("all"))
    except Exception as exc:
      status_lbl.configure(text="STATUS: REPORT ERROR", fg=UI_FAIL)
      print(f"[WARNING] Could not refresh failure-criteria report: {exc}")

  def on_close():
    try:
      win.unbind("<MouseWheel>")
      win.unbind("<Button-4>")
      win.unbind("<Button-5>")
      win.destroy()
    finally:
      if parent_fig is not None:
        parent_fig._failure_criteria_report = None

  # No in-window Close button: use the standard window close control so the
  # report retains the clean report appearance requested by the user.
  win.protocol("WM_DELETE_WINDOW", on_close)

  if parent_fig is not None:
    parent_fig._failure_criteria_report = {"window": win, "refresh": refresh}
  win.update_idletasks()
  _on_content_config()
  refresh()
  return refresh


def launch_unified_dashboard(
    results_store,
    system_flags,
    anchor_db,
    project_path="",
    raster_path="",
):
  """Launch the final unified dashboard after project/raster selection and location picking."""
  _ensure_geo_dependencies()
  fig = plt.figure(figsize=(18, 10), facecolor=UI_BG)
  fig.canvas.mpl_connect("key_press_event", on_key_press)
  try:
    fig.canvas.manager.set_window_title(
        "Nurdins Mooring Solution - Unified Dashboard"
    )
  except Exception:
    pass

  available_systems = [
      sys_name
      for sys_name, flag in system_flags.items()
      if flag and sys_name in results_store
  ]
  if not available_systems:
    print("\n[WARNING] No active systems available to display in unified window.")
    return

  current_state = {
      "system": available_systems[0],
      "anchor": "DEA",
      "project": project_path,
      "raster": raster_path,
      "csv": "",
      "profile_index": 0,
      "turbine_opacity": 1.0,
  }
  current_local_results = []

  for anc in anchor_db.keys():
    if (
        current_state["system"] in results_store
        and anc in results_store[current_state["system"]]
        and results_store[current_state["system"]][anc]["success"]
    ):
      current_state["anchor"] = anc
      break

  # Reserve a distinct title band above the 2-D profile.  The profile title
  # can include the branch, anchor type and local raster depth, so it must not
  # share vertical space with the navigation controls.
  ax1 = fig.add_axes([0.21, 0.28, 0.37, 0.60])
  ax2 = fig.add_axes([0.61, 0.28, 0.37, 0.65], projection="3d")
  # Keep the 3-D opacity control immediately above the export-controls panel.
  ax_turbine_opacity = fig.add_axes([0.665, 0.258, 0.27, 0.022], facecolor=UI_PANEL)
  turbine_opacity_slider = Slider(
      ax_turbine_opacity,
      "3D turbine opacity",
      0.0,
      1.0,
      valinit=float(current_state["turbine_opacity"]),
      valstep=0.05,
      color=UI_ACCENT,
  )

  # 2D profile navigation: cycles T1L1 -> T1L2 -> ... -> T2L1.
  ax_btn_prev_profile = fig.add_axes([0.21, 0.956, 0.11, 0.030])
  btn_prev_profile = Button(ax_btn_prev_profile, "← Previous", color="lightgray", hovercolor="lightblue")
  ax_btn_next_profile = fig.add_axes([0.47, 0.956, 0.11, 0.030])
  btn_next_profile = Button(ax_btn_next_profile, "Next →", color="lightgray", hovercolor="lightblue")
  profile_label = fig.text(0.395, 0.971, "T1L1", ha="center", va="center", fontsize=9, fontweight="bold")

  ax_btn_failure_criteria = fig.add_axes([0.875, 0.956, 0.105, 0.030])
  btn_failure_criteria = Button(ax_btn_failure_criteria, "Failure Criteria", color="lightgray", hovercolor="lightblue")

  ax_wd_panel = fig.add_axes([0.01, 0.88, 0.14, 0.05])
  ax_wd_panel.set_facecolor(UI_PANEL)
  ax_wd_panel.set_xticks([])
  ax_wd_panel.set_yticks([])
  for spine in ax_wd_panel.spines.values():
    spine.set_color("#adb5bd")
    spine.set_linewidth(1.5)

  water_depth_label = fig.text(
      0.08, 0.905, "User Defined Water Depth: -- m\n(Threshold: --)", fontsize=7.5,
      fontweight="bold", color="#2c3e50", ha="center", va="center"
  )

  ax_sidebar = fig.add_axes([0.01, 0.21, 0.14, 0.67])
  ax_sidebar.set_facecolor(UI_PANEL)
  ax_sidebar.set_xticks([])
  ax_sidebar.set_yticks([])
  for spine in ax_sidebar.spines.values():
    spine.set_color("#d0d0d0")

  # Small persistent overview map: raster + selected centre + generated farm
  # perimeter. Keep it centred within the sidebar width and entirely left of
  # the export-controls panel, rather than letting the map span into it.
  _overview_width = 0.13
  _overview_left = 0.01 + (0.14 - _overview_width) / 2.0
  ax_overview = fig.add_axes([_overview_left, 0.015, _overview_width, 0.17])
  ax_overview.set_facecolor(UI_PANEL)
  for spine in ax_overview.spines.values():
    spine.set_color("#adb5bd")
    spine.set_linewidth(1.0)
  ax_overview.set_xticks([])
  ax_overview.set_yticks([])
  ax_overview.set_title("Site Overview", fontsize=8, fontweight="bold", pad=3)

  overview_raster = None
  overview_transform = None
  overview_crs = None
  overview_bounds = None
  try:
    with rasterio.open(raster_path) as src_over: 
      overview_crs = src_over.crs or "EPSG:4326"
      max_dim = 700
      scale = min(1.0, max_dim / max(src_over.width, src_over.height))
      oh = max(1, int(src_over.height * scale))
      ow = max(1, int(src_over.width * scale))
      # Use the same Navia ramp as the main location picker.
      overview_raster = src_over.read(1, out_shape=(oh, ow), masked=True)
      overview_bounds = src_over.bounds
      overview_transform = Transformer.from_crs("EPSG:4326", overview_crs, always_xy=True)
  except Exception as e:
    print(f"[WARNING] Could not load dashboard overview raster: {e}")

  if overview_raster is not None and overview_bounds is not None:
    overview_values = np.asarray(
        overview_raster.compressed() if np.ma.isMaskedArray(overview_raster)
        else overview_raster
    ).ravel()
    overview_values = overview_values[np.isfinite(overview_values)]
    if overview_values.size:
      overview_vmin = float(np.percentile(overview_values, 2.0))
      overview_vmax = float(np.percentile(overview_values, 98.0))
      if overview_vmax <= overview_vmin:
        overview_vmin = float(np.min(overview_values))
        overview_vmax = float(np.max(overview_values))
    else:
      overview_vmin, overview_vmax = None, None

    ax_overview.imshow(
        overview_raster,
        extent=(
            overview_bounds.left, overview_bounds.right,
            overview_bounds.bottom, overview_bounds.top
        ),
        origin="upper",
        cmap=NAVIA_CMAP,
        vmin=overview_vmin,
        vmax=overview_vmax,
    )
    ax_overview.set_xlim(overview_bounds.left, overview_bounds.right)
    ax_overview.set_ylim(overview_bounds.bottom, overview_bounds.top)

  overview_perimeter_plot, = ax_overview.plot([], [], "-", linewidth=1.6, color="purple", zorder=19, label="Farm Perimeter")
  ax_overview.legend(loc="lower right", fontsize=5.5, framealpha=0.8)

  fig.text(0.02, 0.835, "Mooring Systems", fontsize=9.5, fontweight="bold", color="#2c3e50")
  fig.patches.append(patches.Rectangle((0.015, 0.82), 0.125, 0.002, transform=fig.transFigure, color="#bdc3c7", clip_on=False))

  sys_buttons = {}
  sys_labels = {}
  # Five dots in each system button provide an at-a-glance status summary for
  # DEA, Suction Pile, Driven Pile, Drilled Pile, and Gravity respectively.
  # They are updated from the same solver/raster checks as the anchor buttons.
  sys_status_dots = {}
  sys_list = ["Catenary", "Semi-Taut", "Taut"]
  for idx, s_name in enumerate(sys_list):
    ax_s = fig.add_axes([0.02, 0.76 - (idx * 0.05), 0.075, 0.04])
    is_active = system_flags.get(s_name, False) and s_name in results_store
    btn_s = Button(ax_s, s_name, color="lightgray", hovercolor=UI_ACCENT_HOVER if is_active else "lightgray")
    # Keep the system name in the upper portion of its button, leaving room
    # for the five small anchor-status circles along the bottom.
    btn_s.label.set_position((0.5, 0.66))
    btn_s.label.set_fontsize(7.5)
    sys_buttons[s_name] = btn_s
    dots = []
    for dot_idx in range(5):
      # Scatter marker sizes are measured in screen points, so they remain
      # true circles even though a button axis is much wider than it is tall.
      dot = ax_s.scatter(
          [0.14 + dot_idx * 0.18], [0.19],
          marker="o",
          s=24,
          transform=ax_s.transAxes,
          facecolors="#b0b0b0",
          edgecolors="#666666",
          linewidths=0.45,
          zorder=5,
          clip_on=False,
      )
      dots.append(dot)
    sys_status_dots[s_name] = dots
    lbl_s = fig.text(0.10, 0.78 - (idx * 0.05), "", fontsize=7.5, ha="left", va="center", fontweight="bold")
    sys_labels[s_name] = lbl_s

  fig.patches.append(patches.Rectangle((0.015, 0.59), 0.125, 0.002, transform=fig.transFigure, color="#bdc3c7", clip_on=False))
  fig.text(0.02, 0.57, "Anchor Types & Status", fontsize=9.5, fontweight="bold", color="#2c3e50")
  fig.patches.append(patches.Rectangle((0.015, 0.555), 0.125, 0.002, transform=fig.transFigure, color="#bdc3c7", clip_on=False))

  anchor_buttons = {}
  anchor_labels = {}
  anchor_list = list(anchor_db.keys())
  for idx, a_name in enumerate(anchor_list):
    ax_a = fig.add_axes([0.02, 0.49 - (idx * 0.065), 0.075, 0.055])
    btn_a = Button(ax_a, a_name, color="lightgray", hovercolor="limegreen")
    anchor_buttons[a_name] = btn_a
    lbl = fig.text(0.10, 0.517 - (idx * 0.065), "", fontsize=7.5, ha="left", va="center", fontweight="bold")
    anchor_labels[a_name] = lbl

  # Workflow controls now only show the paths selected on the first screen,
  # plus the CSV generated by the Save to CSV button.
  # TextBox labels are drawn just to the left of their axes. Keep a generous
  # left/top margin so the QGIS Project, Raster File, and Layout CSV labels
  # remain visually contained by the export-controls background panel.
  # Lower and compact the export panel so the seismic-volume action has its
  # own clear row above it, without overlapping the 3-D opacity control.
  ax_panel = fig.add_axes([0.15, 0.005, 0.83, 0.195])
  ax_panel.set_facecolor("#f8f9fa")
  ax_panel.set_xticks([])
  ax_panel.set_yticks([])
  for spine in ax_panel.spines.values():
    spine.set_color("#d0d0d0")
  # Keep the map above the dashboard background layer.
  ax_overview.set_zorder(ax_panel.get_zorder() + 1)

  fig.text(0.23, 0.173, "Project / Raster / QGIS Export Controls", fontsize=9.0, fontweight="bold", color="#333333")

  ax_txt_qgis_file = fig.add_axes([0.23, 0.130, 0.67, 0.026])
  txt_qgis_file = TextBox(ax_txt_qgis_file, "QGIS Project: ", initial=project_path or "", textalignment="left")
  ax_txt_raster_file = fig.add_axes([0.23, 0.092, 0.67, 0.026])
  txt_raster_file = TextBox(ax_txt_raster_file, "Raster File: ", initial=raster_path or "", textalignment="left")
  ax_txt_csv_file = fig.add_axes([0.23, 0.054, 0.67, 0.026])
  txt_csv_file = TextBox(ax_txt_csv_file, "Layout CSV: ", initial="", textalignment="left")

  # These paths are intentionally read-only in the final dashboard.
  txt_qgis_file.set_active(False)
  txt_raster_file.set_active(False)
  txt_csv_file.set_active(False)
  txt_qgis_file.text_disp.set_clip_on(True)
  txt_raster_file.text_disp.set_clip_on(True)
  txt_csv_file.text_disp.set_clip_on(True)

  ax_btn_back = fig.add_axes([0.23, 0.012, 0.18, 0.029])
  btn_back = Button(ax_btn_back, "← Location Picker", color="lightgray", hovercolor="lightblue")

  ax_btn_csv = fig.add_axes([0.43, 0.012, 0.23, 0.029])
  btn_export = Button(ax_btn_csv, "Save QGIS CSV", color="lightgray", hovercolor="lightblue")

  ax_btn_qgis = fig.add_axes([0.69, 0.012, 0.21, 0.029])
  btn_qgis = Button(ax_btn_qgis, "Open in Desktop QGIS", color="lightgray", hovercolor="lightgray")

  ax_btn_seismic = fig.add_axes([0.23, 0.214, 0.67, 0.035])
  btn_seismic = Button(
      ax_btn_seismic, "Seismic Volume Calculation",
      color="lightgray", hovercolor="lightblue",
  )

  # Keep widget objects attached to the Figure.  This prevents accidental
  # garbage collection and makes the callbacks reliable on TkAgg/Windows.
  fig._dashboard_widgets = [
      *sys_buttons.values(),
      *anchor_buttons.values(),
      txt_qgis_file, txt_raster_file, txt_csv_file,
      btn_export, btn_qgis, btn_back, btn_seismic, btn_prev_profile, btn_next_profile,
      btn_failure_criteria, turbine_opacity_slider,
  ]
  for _btn_ax in (ax_btn_back, ax_btn_csv, ax_btn_qgis, ax_btn_seismic, ax_btn_prev_profile, ax_btn_next_profile, ax_btn_failure_criteria, ax_turbine_opacity):
    _btn_ax.set_zorder(100)
  for _btn in (*sys_buttons.values(), *anchor_buttons.values()):
    _btn.ax.set_zorder(100)

  def get_failure_report_payload():
    sys_name = current_state.get("system", "--")
    anc_name = current_state.get("anchor", "--")
    data = results_store.get(sys_name, {}).get(anc_name, {})
    nlines = max(1, int(data.get("num_lines", 1)))
    nturb = max(1, int(data.get("num_turbines", 1)))
    pi = int(current_state.get("profile_index", 0))
    if bool(data.get("inverted_bridle", False)):
      base_pi = (pi // 2) % max(1, nturb * nlines)
      profile = f"T{base_pi // nlines + 1}L{base_pi % nlines + 1}{'a' if pi % 2 == 0 else 'b'}"
    else:
      base_pi = pi % max(1, nturb * nlines)
      profile = f"T{base_pi // nlines + 1}L{base_pi % nlines + 1}"
    selected = current_local_results[0] if current_local_results else {}
    # Raster-selected profile values override the base result where available.
    if selected:
      merged = dict(data)
      for key in (
          "local_depth", "success", "status", "msg", "ground_ratio",
          "grounded_surface_length", "suspended_surface_length",
          "suspended_line_angle_deg", "min_geometric_length", "direct_geometric_span", "excess_line",
          "T_fairlead", "T_padeye", "T_max", "HF", "VF", "HA", "VA",
          "taut_min_tension", "taut_max_tension", "line_fos",
          "line_fos_chain", "line_fos_rope", "line_fos_fairlead",
          "line_fos_padeye", "chain_grade", "rope_material",
          "_raster_system_pass", "_raster_failure_reasons", "_failure_stage",
      ):
        if key in selected and selected.get(key) is not None:
          merged[key] = selected.get(key)
      # Preserve result-level data such as diameters and anchor results.
      data_for_report = merged
    else:
      data_for_report = data
    # ------------------------------------------------------------------
    # Inverted-bridle report selection
    # ------------------------------------------------------------------
    # The normal solver result contains the common fairlead/trunk loads, but
    # the selected T1L1a/T1L1b profile has its OWN padeye H/V load and anchor
    # capacity.  The 2-D information box already uses these branch values;
    # the Failure Criteria report must use the same branch for its anchor/
    # padeye criteria as well.
    if bool(data_for_report.get("inverted_bridle", False)) and (pi % 2 in (0, 1)):
      branch_no = 1 if (pi % 2 == 0) else 2
      branch_loads = (data_for_report.get("bridle_loads") or {}).get("branches", [])
      branch_load = branch_loads[branch_no - 1] if len(branch_loads) >= branch_no else {}

      # Prefer branch data from the raster-selected result, then the base result.
      if not branch_load and selected:
        branch_loads = (selected.get("bridle_loads") or {}).get("branches", [])
        branch_load = branch_loads[branch_no - 1] if len(branch_loads) >= branch_no else {}

      if branch_load:
        branch_H = _safe_float(branch_load.get("HA", np.nan))
        branch_V = _safe_float(branch_load.get("VA", np.nan))
        branch_T = _safe_float(branch_load.get("tension", np.nan))

        if np.isfinite(branch_H):
          data_for_report["HA"] = branch_H
        if np.isfinite(branch_V):
          data_for_report["VA"] = branch_V
        if np.isfinite(branch_T):
          data_for_report["T_padeye"] = branch_T
          # Keep the common fairlead load, but the selected branch can govern
          # the overall line tension shown by the report.
          tf = _safe_float(data_for_report.get("T_fairlead", np.nan))
          data_for_report["T_max"] = max(v for v in (branch_T, tf) if np.isfinite(v)) if (np.isfinite(branch_T) or np.isfinite(tf)) else np.nan

          # Selected-branch line FoS: use the weakest physical line component
          # in that bridle branch. Catenary uses chain only; Semi-Taut has both
          # chain and synthetic rope; Taut uses the synthetic rope.
          try:
            chain_d_report = _safe_float(data_for_report.get("chain_d", np.nan))
            rope_d_report = _safe_float(data_for_report.get("rope_d", np.nan))
            mb_ls = []
            if str(sys_name).lower() != "taut" and np.isfinite(chain_d_report) and chain_d_report > 0:
              mb_ls.append(calculate_chain_mbl(chain_d_report, str(data_for_report.get("chain_grade", "R4"))))
            if np.isfinite(rope_d_report) and rope_d_report > 0:
              mat = str(data_for_report.get("rope_material", "Polyester" if str(sys_name).lower() == "semi-taut" else "HMPE"))
              mb_ls.append(calculate_rope_mbl(rope_d_report, mat))
            if mb_ls and np.isfinite(branch_T) and branch_T > 0:
              branch_fos = min(mb_ls) / branch_T
              data_for_report["line_fos_padeye"] = branch_fos
              data_for_report["line_fos"] = min(_safe_float(data_for_report.get("line_fos_fairlead", np.inf)), branch_fos)
              if str(sys_name).lower() == "semi-taut" and data_for_report.get("line_fos_rope") is not None:
                data_for_report["line_fos_rope"] = min(_safe_float(data_for_report.get("line_fos_rope", np.inf)), branch_fos)
              if str(sys_name).lower() == "semi-taut" and data_for_report.get("line_fos_chain") is not None:
                data_for_report["line_fos_chain"] = min(_safe_float(data_for_report.get("line_fos_chain", np.inf)), branch_fos)
          except Exception:
            pass

        # Branch-specific anchor capacity is authoritative for the selected
        # physical padeye.  Never fall back to the centreline anchor result
        # when the selected bridle branch has a calculated result.
        branch_anchor = (data_for_report.get("bridle_anchor_results") or {}).get(f"Branch {branch_no}", {})
        if not branch_anchor and selected:
          branch_anchor = (selected.get("bridle_anchor_results") or {}).get(f"Branch {branch_no}", {})
        if isinstance(branch_anchor, dict) and branch_anchor:
          data_for_report["anchor_results"] = {anc_name: dict(branch_anchor)}
        data_for_report["selected_bridle_branch"] = branch_no
        data_for_report["selected_bridle_profile"] = profile

    # For failed configurations, reconstruct any diagnostics that can still be
    # calculated from the stored inputs so the report shows the failed case on
    # the same visual criteria bars rather than replacing them with blank panels.
    data_for_report = enrich_failure_report_data(data_for_report, sys_name, anchor_db)
    # enrich_failure_report_data deliberately preserves branch-resolved values;
    # re-apply the anchor result afterwards if it was branch-specific.
    if bool(data_for_report.get("inverted_bridle", False)) and data_for_report.get("selected_bridle_branch"):
      bno = int(data_for_report["selected_bridle_branch"])
      banch = (data_for_report.get("bridle_anchor_results") or {}).get(f"Branch {bno}", {})
      if isinstance(banch, dict) and banch:
        data_for_report["anchor_results"] = {anc_name: dict(banch)}
    data_for_report["_failure_stage"] = infer_failure_stage(sys_name, data_for_report)
    return {"system": sys_name, "anchor": anc_name, "profile": profile, "data": data_for_report}

  def refresh_failure_criteria_report():
    report = getattr(fig, "_failure_criteria_report", None)
    if report is not None:
      try:
        report["refresh"]()
      except Exception as exc:
        print(f"[WARNING] Could not update failure criteria report: {exc}")

  def update_qgis_button_states():
    proj = current_state["project"].strip()
    csv_f = current_state["csv"].strip()
    ready = bool(proj and csv_f and os.path.exists(proj) and os.path.exists(csv_f))
    btn_qgis.color = "lightgreen" if ready else "lightgray"
    btn_qgis.hovercolor = "limegreen" if ready else "lightgray"
    btn_qgis.ax.set_facecolor(btn_qgis.color)
    fig.canvas.draw_idle()

  def redraw_plots():
    sys_name = current_state["system"]
    anc_name = current_state["anchor"]
    if (
        sys_name not in results_store
        or anc_name not in results_store[sys_name]
        or not results_store[sys_name][anc_name]["success"]
    ):
      ax1.clear()
      ax2.clear()
      ax1.text(0.5, 0.5, f"Configuration '{sys_name} + {anc_name}' not available or failed.",
               ha="center", va="center", transform=ax1.transAxes, fontsize=10,
               fontweight="bold", color="red")
      ax2.set_title(f"Farm Layout - {sys_name} ({anc_name}) [FAILED]")
      refresh_failure_criteria_report()
      fig.canvas.draw_idle()
      return
    data = results_store[sys_name][anc_name]
    num_lines_local = max(1, int(data.get("num_lines", 1)))
    num_turbines_local = max(1, int(data.get("num_turbines", 1)))
    bridle_navigation = bool(data.get("inverted_bridle", False))
    total_profiles = max(1, num_turbines_local * num_lines_local * (2 if bridle_navigation else 1))
    current_state["profile_index"] %= total_profiles
    pi = current_state["profile_index"]
    if bridle_navigation:
      base_pi = pi // 2
      profile_label.set_text(f"T{base_pi // num_lines_local + 1}L{base_pi % num_lines_local + 1}{'a' if pi % 2 == 0 else 'b'}")
    else:
      profile_label.set_text(f"T{pi // num_lines_local + 1}L{pi % num_lines_local + 1}")
    current_local_results[:] = render_terrain_dashboard(
        ax1, ax2, data, raster_path, anchor_db, anc_name,
        profile_index=pi,
        turbine_opacity=current_state["turbine_opacity"],
    )
    refresh_failure_criteria_report()
    # Update PASS/FAIL controls immediately from the exact raster rendering
    # result just calculated.  This prevents a failed raster turbine from
    # remaining selectable in the left-hand dashboard.
    update_ui_states()
    fig.canvas.draw_idle()

  def update_turbine_opacity(value):
    """Redraw the dashboard with the requested 3-D turbine/mooring opacity."""
    opacity = float(np.clip(value, 0.0, 1.0))
    if abs(opacity - float(current_state["turbine_opacity"])) < 1e-9:
      return
    current_state["turbine_opacity"] = opacity
    redraw_plots()

  fig._turbine_opacity_callback = turbine_opacity_slider.on_changed(update_turbine_opacity)

  def update_overview_map():
    """Refresh the small raster/site overview whenever the active mooring configuration changes."""
    try:
      data = results_store[current_state["system"]][current_state["anchor"]]
      if not data.get("success"):
        overview_perimeter_plot.set_data([], [])
        fig.canvas.draw_idle()
        return

      # Use the same lightweight perimeter helper as the CSV export.  Building
      # every mooring/radius CSV row here just to obtain 120 perimeter points
      # was unnecessarily expensive whenever a selection changed.
      ox, oy = latlon_to_epsg3857(data["center_lat"], data["center_lon"])
      perimeter_x, perimeter_y = farm_perimeter_coordinates(data, ox, oy)
      if perimeter_x.size:
        tx = Transformer.from_crs("EPSG:3857", overview_crs, always_xy=True)
        px, py = tx.transform(perimeter_x, perimeter_y)
        px, py = np.asarray(px, dtype=float), np.asarray(py, dtype=float)
        overview_perimeter_plot.set_data(px, py)

        # Keep the minimap on the FULL raster extent.  No centre marker is
        # displayed; only the farm perimeter is overlaid.
        ax_overview.set_xlim(overview_bounds.left, overview_bounds.right)
        ax_overview.set_ylim(overview_bounds.bottom, overview_bounds.top)
      else:
        overview_perimeter_plot.set_data([], [])
        ax_overview.set_xlim(overview_bounds.left, overview_bounds.right)
        ax_overview.set_ylim(overview_bounds.bottom, overview_bounds.top)
      fig.canvas.draw_idle()
    except Exception as e:
      print(f"[WARNING] Could not update dashboard overview map: {e}")

  def get_anchor_status(sys_name, anc_name):
    """Return the dashboard status and colour for one system/anchor pair."""
    data = results_store.get(sys_name, {}).get(anc_name)
    if not data:
      return "N/A", "#b0b0b0"

    base_pass = bool(data.get("success"))
    if not base_pass:
      return "FAIL", "#e74c3c"

    raster_checked = "_raster_system_pass" in data
    raster_pass = bool(data.get("_raster_system_pass")) if raster_checked else base_pass
    turbine_states = [
        bool(item.get("success", False))
        for item in data.get("_raster_turbine_results", [])
    ]
    if raster_checked and turbine_states and any(turbine_states) and not all(turbine_states):
      return "WARNING", "#f39c12"
    if raster_checked and not raster_pass:
      return "FAIL", "#e74c3c"
    return "PASS", "#2ecc71"

  def update_ui_states():
    sys_name = current_state["system"]
    anc_name = current_state["anchor"]
    _selected_data = results_store.get(sys_name, {}).get(anc_name, {})
    _design_depth = abs(_safe_float(_selected_data.get("water_depth", np.nan)))
    _threshold_fraction = normalise_raster_depth_threshold(
        _selected_data.get(
            "raster_depth_threshold", RASTER_DEPTH_TOLERANCE_FRACTION
        )
    )
    if _threshold_fraction is None:
      _threshold_fraction = RASTER_DEPTH_TOLERANCE_FRACTION
    if np.isfinite(_design_depth) and _design_depth > 0.0:
      # This panel reports the Excel design input and H30 tolerance, not the
      # first turbine's sampled raster depth. Local depths remain available in
      # the selected profile and raster-depth suitability report.
      water_depth_label.set_text(
          f"User Defined Water Depth: {_design_depth:.1f} m\n"
          f"(Threshold: +/-{float(_threshold_fraction) * 100.0:.2f}%)"
      )
    else:
      water_depth_label.set_text("User Defined Water Depth: -- m\n(Threshold: --)")

    # Anchor buttons: preserve the original PASS/FAIL logic, but once the
    # raster has actually been rendered for this configuration, use the
    # raster turbine-by-turbine viability as the authoritative UI status.
    for anc_name in anchor_list:
      btn = anchor_buttons[anc_name]
      lbl = anchor_labels[anc_name]
      base_exists = sys_name in results_store and anc_name in results_store[sys_name]
      base_pass = bool(results_store[sys_name][anc_name].get("success")) if base_exists else False
      raster_checked = base_exists and "_raster_system_pass" in results_store[sys_name][anc_name]
      raster_pass = bool(results_store[sys_name][anc_name].get("_raster_system_pass")) if raster_checked else base_pass
      turbine_results = (
          results_store[sys_name][anc_name].get("_raster_turbine_results", [])
          if base_exists else []
      )
      turbine_states = [bool(item.get("success", False)) for item in turbine_results]
      mixed_turbine_result = (
          raster_checked
          and bool(turbine_states)
          and any(turbine_states)
          and not all(turbine_states)
      )

      if not base_exists:
        btn.color = "lightgray"
        btn.hovercolor = "lightgray"
        btn.ax.set_facecolor("lightgray")
        lbl.set_text("N/A")
        lbl.set_color("gray")
        btn.set_active(False)
      elif not base_pass:
        btn.color = "#f5b7b1" if anc_name != current_state["anchor"] else "#ec7063"
        btn.hovercolor = "#e74c3c"
        btn.ax.set_facecolor(btn.color)
        lbl.set_text("FAIL")
        lbl.set_color("red")
        btn.set_active(True)
      elif mixed_turbine_result:
        btn.color = "#f8c471" if anc_name != current_state["anchor"] else "#f5b041"
        btn.hovercolor = "#e67e22"
        btn.ax.set_facecolor(btn.color)
        lbl.set_text("WARNING")
        lbl.set_color("#b66b00")
        btn.set_active(True)
      elif raster_checked and not raster_pass:
        btn.color = "#f5b7b1" if anc_name != current_state["anchor"] else "#ec7063"
        btn.hovercolor = "#e74c3c"
        btn.ax.set_facecolor(btn.color)
        lbl.set_text("FAIL")
        lbl.set_color("red")
        btn.set_active(True)
      else:
        btn.color = "lightgreen" if anc_name == current_state["anchor"] else "lightgray"
        btn.hovercolor = "limegreen"
        btn.ax.set_facecolor(btn.color)
        lbl.set_text("PASS")
        lbl.set_color("green")
        btn.set_active(True)

    # System buttons remain navigation controls, with five colour dots that
    # mirror the PASS/WARNING/FAIL status of the five anchor types below.
    for s_name, btn in sys_buttons.items():
      lbl_s = sys_labels[s_name]
      is_enabled = system_flags.get(s_name, False) and s_name in results_store
      for dot_idx, dot in enumerate(sys_status_dots[s_name]):
        if is_enabled and dot_idx < len(anchor_list):
          _, dot_colour = get_anchor_status(s_name, anchor_list[dot_idx])
        else:
          dot_colour = "#b0b0b0"
        dot.set_facecolor(dot_colour)
        dot.set_edgecolor("#666666")
      if not is_enabled:
        lbl_s.set_text("N/A")
        lbl_s.set_color("gray")
        btn.color = "lightgray"
        btn.hovercolor = "lightgray"
        btn.ax.set_facecolor("lightgray")
        btn.set_active(False)
      else:
        lbl_s.set_text("")
        btn.color = "deepskyblue" if s_name == sys_name else "lightgray"
        btn.hovercolor = "deepskyblue"
        btn.ax.set_facecolor(btn.color)
        btn.set_active(True)

  def select_system(s_name):
    if not system_flags.get(s_name, False) or s_name not in results_store:
      return

    # Whenever a new mooring system is selected, raster-test EVERY anchor type
    # before choosing the first selectable anchor.  This makes the sidebar
    # statuses authoritative immediately rather than testing anchors lazily.
    for anc in anchor_list:
      if anc in results_store[s_name] and results_store[s_name][anc].get("success"):
        evaluate_raster_configuration(results_store[s_name][anc], raster_path, anchor_db, anc)

    current_state["system"] = s_name
    current_state["profile_index"] = 0
    selected_anchor = None
    for anc in anchor_list:
      if (anc in results_store[s_name] and results_store[s_name][anc].get("success")
          and results_store[s_name][anc].get("_raster_system_pass", False)):
        selected_anchor = anc
        break
    if selected_anchor is None:
      for anc in anchor_list:
        if anc in results_store[s_name]:
          selected_anchor = anc
          break
    if selected_anchor is not None:
      current_state["anchor"] = selected_anchor
    redraw_plots()
    update_ui_states()
    update_overview_map()

  def select_anchor(anc_name):
    sys_name = current_state["system"]
    data = results_store.get(sys_name, {}).get(anc_name)
    if not data:
      return
    base_pass = bool(data.get("success"))
    raster_checked = "_raster_system_pass" in data
    raster_pass = bool(data.get("_raster_system_pass")) if raster_checked else base_pass
    # H30-only failures have valid solver geometry.  Let users inspect that
    # geometry (rendered red) while the dashboard status still correctly reads
    # FAIL.  Solver failures retain the existing report-only behaviour.
    raster_depth_only_failure = (
        raster_checked
        and not raster_pass
        and bool(data.get("_raster_depth_band_failure", False))
        and bool(data.get("_raster_solver_system_pass", False))
    )
    if not base_pass or (raster_checked and not raster_pass and not raster_depth_only_failure):
      # Failed anchors are now inspected through the live Failure Criteria report.
      # Selecting a failed anchor changes the active report configuration instead
      # of opening a separate message box.
      current_state["anchor"] = anc_name
      current_state["profile_index"] = 0
      current_local_results[:] = []
      redraw_plots()
      update_ui_states()
      update_overview_map()
      report = getattr(fig, "_failure_criteria_report", None)
      if report is None:
        show_failure_criteria_window(fig, get_failure_report_payload)
      else:
        try:
          report["refresh"]()
          report["window"].deiconify()
          report["window"].lift()
          report["window"].focus_force()
        except Exception:
          show_failure_criteria_window(fig, get_failure_report_payload)
      return
    current_state["anchor"] = anc_name
    current_state["profile_index"] = 0
    redraw_plots()
    update_ui_states()
    update_overview_map()

  for s_name, btn in sys_buttons.items():
    btn.on_clicked(lambda event, sn=s_name: select_system(sn))
  for anc_name, btn in anchor_buttons.items():
    btn.on_clicked(lambda event, an=anc_name: select_anchor(an))

  def on_failure_criteria_clicked(event=None):
    show_failure_criteria_window(fig, get_failure_report_payload)

  btn_failure_criteria.on_clicked(on_failure_criteria_clicked)

  def on_back_clicked(event=None):
    current_lat = None
    current_lon = None
    current_depth = 100.0
    current_depth_threshold = RASTER_DEPTH_TOLERANCE_FRACTION
    try:
      # Preserve the current centre while returning to the point picker.
      d = results_store[current_state["system"]][current_state["anchor"]]
      current_lat = d.get("center_lat")
      current_lon = d.get("center_lon")
      current_depth = d.get("water_depth", current_depth)
      current_depth_threshold = d.get(
          "raster_depth_threshold", current_depth_threshold
      )
    except Exception:
      pass
    plt.close(fig)
    open_raster_location_picker(
        current_state["raster"],
        lambda lat, lon: (
            update_result_locations(results_store, lat, lon),
            launch_unified_dashboard(
                results_store, system_flags, anchor_db,
                project_path=current_state["project"],
                raster_path=current_state["raster"],
            ),
        ),
        water_depth=current_depth,
        depth_threshold=current_depth_threshold,
        initial_lat=current_lat,
        initial_lon=current_lon,
        on_back=lambda: launch_project_setup_screen(
            results_store, system_flags, anchor_db,
            project_path=current_state["project"],
            raster_path=current_state["raster"],
            water_depth=current_depth,
            depth_threshold=current_depth_threshold,
        ),
    )

  def change_profile(delta):
    sys_name = current_state["system"]
    anc_name = current_state["anchor"]
    data = results_store.get(sys_name, {}).get(anc_name)
    if not data:
      return
    total = max(1, int(data.get("num_turbines", 1)) * int(data.get("num_lines", 1)) * (2 if bool(data.get("inverted_bridle", False)) else 1))
    current_state["profile_index"] = (current_state["profile_index"] + delta) % total
    redraw_plots()

  btn_prev_profile.on_clicked(lambda event: change_profile(-1))
  btn_next_profile.on_clicked(lambda event: change_profile(1))

  def on_export_clicked(event):
    sys_name = current_state["system"]
    anc_name = current_state["anchor"]
    if (
        sys_name not in results_store
        or anc_name not in results_store[sys_name]
        or not results_store[sys_name][anc_name]["success"]
    ):
      return
    data = results_store[sys_name][anc_name]
    origin_x, origin_y = latlon_to_epsg3857(data["center_lat"], data["center_lon"])
    export_payload = {
        "system_type": data["system_type"],
        "primary_anc": data["primary_anc"],
        "rows": generate_export_rows(data, origin_x, origin_y),
    }
    saved_file = export_system_to_csv(export_payload, parent_figure=fig)
    if saved_file:
      current_state["csv"] = saved_file
      txt_csv_file.set_val(saved_file)
      update_qgis_button_states()

  def on_qgis_clicked(event):
    proj = current_state["project"].strip()
    csv_file = current_state["csv"].strip()
    if proj and csv_file and os.path.exists(proj) and os.path.exists(csv_file):
      open_qgis_and_import(csv_file, project_path=proj)
    else:
      print("\n[WARNING] Save the QGIS CSV before opening the project in QGIS.")

  def on_seismic_volume_clicked(event):
    """Open the Main.py-native PyVista seismic screen for the active result."""
    sys_name = current_state["system"]
    anc_name = current_state["anchor"]
    raster_file = current_state["raster"].strip()
    data = results_store.get(sys_name, {}).get(anc_name)
    if not data:
      print("[WARNING] No active mooring result is available for Seismic Volume Calculation.")
      return
    if not raster_file or not os.path.isfile(raster_file):
      print("[WARNING] Select a valid raster file before opening Seismic Volume Calculation.")
      return
    try:
      # The evaluation is cached when the active raster/result is unchanged.
      # Refreshing it here ensures the in-memory export contains the same
      # terrain-resolved Taut and bridle geometry shown by the dashboard.
      evaluate_raster_configuration(data, raster_file, anchor_db, anc_name)
      origin_x, origin_y = latlon_to_epsg3857(data["center_lat"], data["center_lon"])
      rows = generate_export_rows(data, origin_x, origin_y)
    except Exception as exc:
      print(f"[ERROR] Could not prepare current mooring geometry for seismic calculation: {exc}")
      return
    try:
      launch_seismic_volume_calculation_detached(data, raster_file, rows)
      print("[INFO] Opening Seismic Volume Calculation in a separate VTK window.")
    except Exception as exc:
      print(f"[ERROR] Could not start Seismic Volume Calculation: {exc}")

  btn_export.on_clicked(on_export_clicked)
  btn_qgis.on_clicked(on_qgis_clicked)
  btn_seismic.on_clicked(on_seismic_volume_clicked)
  btn_back.on_clicked(on_back_clicked)

  # Test all anchor types for the initially selected mooring system before
  # displaying the sidebar statuses.
  for anc in anchor_list:
    if anc in results_store[current_state["system"]] and results_store[current_state["system"]][anc].get("success"):
      evaluate_raster_configuration(results_store[current_state["system"]][anc], raster_path, anchor_db, anc)

  update_ui_states()
  update_overview_map()
  update_qgis_button_states()
  redraw_plots()
  # Force an initial draw before entering the GUI event loop so all widget
  # hitboxes are registered immediately.
  fig.canvas.draw()

  try:
    manager = plt.get_current_backend_manager()
    manager.window.state("zoomed")
  except Exception:
    try:
      fig.canvas.manager.window.state("zoomed")
    except Exception:
      pass

  plt.show()


# =========================================================================
# RENDER HELPER FOR 2D & 3D AXES CONTENT IN UNIFIED WINDOW
# =========================================================================
def determine_farm_area(
    sheet, num_turbines, anchor_radius, h5_farm_area, buffer_zone
):
  try:
    f5_val = sheet.range("F5").value
    if isinstance(f5_val, str):
      use_h5 = f5_val.strip().upper() == "TRUE"
    elif f5_val is None:
      use_h5 = True
    else:
      use_h5 = bool(f5_val)
  except Exception:
    use_h5 = True

  if use_h5 and h5_farm_area is not None and h5_farm_area > 0:
    return h5_farm_area
  else:
    min_spacing = (2.0 * anchor_radius) + buffer_zone
    rows, cols = get_optimal_grid_dimensions(num_turbines)
    total_width_x = cols * min_spacing
    total_height_y = rows * min_spacing * 0.866
    optimized_area = (total_width_x * total_height_y) / 1e6
    return max(optimized_area, 0.01)


def read_from_file_catenary(sheet=None, common_params=None):
  try:
    if sheet is None:
      sheet = get_excel_sheet("Mooring_Python")
    if common_params is None:
      common_params = read_common_excel_params(sheet)
    (
        water_depth,
        num_turbines,
        farm_area_input,
        buffer_zone,
        center_lat,
        center_lon,
    ) = common_params
    anchor_radius = excel_cell(sheet, "catenary_anchor_radius", default=749.0, cast_type=float)
    line_length = excel_cell(sheet, "catenary_line_length", default=770.0, cast_type=float)
    raw_chain_d = excel_cell(sheet, "catenary_chain_diameter", default=0.169, cast_type=float)
    num_lines = excel_cell(sheet, "catenary_number_of_lines", default=3, cast_type=int)
    fairlead_draft = excel_cell(sheet, "fairlead_depth", default=-13.5, cast_type=float)

    farm_area = determine_farm_area(
        sheet, num_turbines, anchor_radius, farm_area_input, buffer_zone
    )
    chain_d = raw_chain_d / 1000.0 if raw_chain_d > 1.0 else raw_chain_d
    return (
        water_depth,
        anchor_radius,
        line_length,
        chain_d,
        15.0,
        fairlead_draft,
        num_turbines,
        farm_area,
        num_lines,
        center_lat,
        center_lon,
        buffer_zone,
    )
  except Exception as e:
    print(f"[ERROR] Failed to read Catenary parameters from Excel: {e}")
    return None


def read_from_file_semi_taut(sheet=None, common_params=None):
  try:
    if sheet is None:
      sheet = get_excel_sheet("Mooring_Python")
    if common_params is None:
      common_params = read_common_excel_params(sheet)
    (
        water_depth,
        num_turbines,
        farm_area_input,
        buffer_zone,
        center_lat,
        center_lon,
    ) = common_params
    anchor_radius = excel_cell(sheet, "semi_taut_anchor_radius", default=450.0, cast_type=float)
    line_length = excel_cell(sheet, "semi_taut_line_length", default=400.0, cast_type=float)
    raw_rope_d = excel_cell(sheet, "semi_taut_rope_diameter", default=0.160, cast_type=float)
    raw_chain_d = excel_cell(sheet, "semi_taut_chain_diameter", default=0.120, cast_type=float)
    num_lines = excel_cell(sheet, "semi_taut_number_of_lines", default=3, cast_type=int)
    taut_percentage = excel_cell(sheet, "taut_percentage", default=30.0, cast_type=float)
    fairlead_draft = excel_cell(sheet, "fairlead_depth", default=-13.5, cast_type=float)

    farm_area = determine_farm_area(
        sheet, num_turbines, anchor_radius, farm_area_input, buffer_zone
    )
    return (
        water_depth,
        anchor_radius,
        line_length,
        (raw_rope_d / 1000.0 if raw_rope_d > 1.0 else raw_rope_d),
        (raw_chain_d / 1000.0 if raw_chain_d > 1.0 else raw_chain_d),
        15.0,
        fairlead_draft,
        num_turbines,
        farm_area,
        num_lines,
        center_lat,
        center_lon,
        buffer_zone,
        taut_percentage,
    )
  except Exception as e:
    print(f"[ERROR] Failed to read Semi-Taut parameters from Excel: {e}")
    return None


def read_from_file_taut(sheet=None, common_params=None):
  try:
    if sheet is None:
      sheet = get_excel_sheet("Mooring_Python")
    if common_params is None:
      common_params = read_common_excel_params(sheet)
    (
        water_depth,
        num_turbines,
        farm_area_input,
        buffer_zone,
        center_lat,
        center_lon,
    ) = common_params
    anchor_radius = excel_cell(sheet, "taut_anchor_radius", default=200.0, cast_type=float)
    raw_line_d = excel_cell(sheet, "taut_line_diameter", default=0.140, cast_type=float)
    num_lines = excel_cell(sheet, "taut_number_of_lines", default=3, cast_type=int)
    fairlead_draft = excel_cell(sheet, "fairlead_depth", default=-13.5, cast_type=float)

    farm_area = determine_farm_area(
        sheet, num_turbines, anchor_radius, farm_area_input, buffer_zone
    )
    return (
        water_depth,
        anchor_radius,
        (raw_line_d / 1000.0 if raw_line_d > 1.0 else raw_line_d),
        15.0,
        fairlead_draft,
        num_turbines,
        farm_area,
        num_lines,
        center_lat,
        center_lon,
        buffer_zone,
    )
  except Exception as e:
    print(f"[ERROR] Failed to read Taut parameters from Excel: {e}")
    return None


# =========================================================================
# MAIN EXECUTION
# =========================================================================def epsg3857_to_latlon(x, y):
  """Converts EPSG:3857 (meters) to WGS84 Latitude and Longitude (degrees)."""
  R = 6378137.0
  lon = np.degrees(x / R)
  lat = np.degrees(2.0 * np.arctan(np.exp(y / R)) - np.pi / 2.0)
  return lat, lon




def get_optimal_grid_dimensions(num_turbines):
  """Calculates optimal rows and columns for a given number of turbines to form a balanced, uniform rectangle with equal items per row where possible."""
  if num_turbines <= 0:
    return 1, 1
  divisors = [i for i in range(1, num_turbines + 1) if num_turbines % i == 0]
  target = np.sqrt(num_turbines)
  best_cols = min(divisors, key=lambda x: abs(x - target))
  best_rows = num_turbines // best_cols
  return best_rows, best_cols




def find_qgis_executable():
  """Find QGIS on Windows, macOS, or a PATH-based installation."""
  if sys.platform == "darwin":
    mac_patterns = [
        "/Applications/QGIS*.app/Contents/MacOS/QGIS",
        os.path.expanduser("~/Applications/QGIS*.app/Contents/MacOS/QGIS"),
    ]
    for pattern in mac_patterns:
      matches = glob.glob(pattern)
      if matches:
        return sorted(matches, reverse=True)[0]

  search_patterns = [
      r"C:\Program Files\QGIS*\bin\qgis-ltr-bin.exe",
      r"C:\Program Files\QGIS*\bin\qgis-bin.exe",
      r"C:\OSGeo4W*\bin\qgis-ltr-bin.exe",
      r"C:\OSGeo4W*\bin\qgis-bin.exe",
      r"C:\Program Files\QGIS*\bin\qgis-ltr.bat",
      r"C:\Program Files\QGIS*\bin\qgis.bat",
      r"C:\OSGeo4W*\bin\qgis-ltr.bat",
      r"C:\OSGeo4W*\bin\qgis.bat",
  ]

  for pattern in search_patterns:
    matches = glob.glob(pattern)
    if matches:
      return sorted(matches, reverse=True)[0]

  for exe_name in [
      "qgis-ltr-bin.exe",
      "qgis-bin.exe",
      "qgis-ltr.bat",
      "qgis.bat",
      "qgis",
      "QGIS",
  ]:
    path_in_env = shutil.which(exe_name)
    if path_in_env:
      return path_in_env

  return None




def farm_turbine_coordinates(data):
  """Return the same turbine coordinates used by the existing dashboard/export."""
  origin_x, origin_y = latlon_to_epsg3857(data["center_lat"], data["center_lon"])
  anchor_radius = float(data.get("anchor_radius", 0.0))
  buffer_zone = float(data.get("buffer_zone", 50.0))
  num_turbines = int(data.get("num_turbines", 1))
  farm_area = data.get("farm_area")
  min_spacing = (2.0 * anchor_radius) + buffer_zone
  base_spacing = (
      np.sqrt((farm_area * 1e6) / num_turbines)
      if farm_area and farm_area > 0 else min_spacing
  )
  spacing = max(min_spacing, base_spacing)
  rows, cols = get_optimal_grid_dimensions(num_turbines)
  total_width_x = cols * spacing
  total_height_y = rows * spacing * 0.866
  coords = []
  for i in range(num_turbines):
    row_idx, col_idx = i // cols, i % cols
    cx = origin_x - total_width_x / 2.0 + (col_idx + 0.5 * (row_idx % 2)) * spacing
    cy = origin_y - total_height_y / 2.0 + row_idx * spacing * 0.866
    coords.append((float(cx), float(cy)))
  return coords





def build_mooring_info_lines(data, profile_label=None):
  """Build the reorganized on-screen mooring information panel."""
  system_type = data.get("system_type", "Mooring")
  primary_anc = data.get("primary_anc", "--")
  line_length = float(data.get("line_length", 0.0) or 0.0)
  anchor_radius = float(data.get("anchor_radius", 0.0) or 0.0)
  num_lines = int(data.get("num_lines", 0) or 0)
  num_turbines = int(data.get("num_turbines", 0) or 0)
  raster_depth = float(data.get("water_depth", np.nan))
  fairlead_depth = float(data.get("fairlead_draft", np.nan))
  anchor_width = float(data.get("anchor_width", 0.0) or 0.0)
  anchor_height = float(data.get("anchor_height", 0.0) or 0.0)
  anchor_depth = float(data.get("anchor_depth", 0.0) or 0.0)
  L_sub = float(data.get("L_sub", 0.0) or 0.0)
  L_bot = float(data.get("grounded_surface_length", data.get("L_bot", 0.0)) or 0.0)
  L_sus = float(data.get("suspended_surface_length", data.get("L_sus", 0.0)) or 0.0)
  min_geom = float(data.get("min_geometric_length", np.nan))
  excess = float(data.get("excess_line", np.nan))
  ground_ratio = float(data.get("ground_ratio", 0.0) or 0.0)
  padeye_cfg = (data.get("padeye_params") or {}).get(primary_anc, {})
  padeye_angle = padeye_cfg.get("angle_deg", data.get("padeye_angle_deg"))
  padeye_fraction = padeye_cfg.get("position_fraction", data.get("padeye_position_fraction"))
  padeye_z = data.get("padeye_z", data.get("z_anchor_custom"))
  try:
    padeye_z = float(padeye_z)
  except (TypeError, ValueError):
    padeye_z = np.nan
  # A selected Taut bridle branch supplies its own raster-derived padeye Z.
  # Fall back to the local line depth only when no explicit physical endpoint
  # has been provided (the regular Taut-system case).
  if system_type == "Taut" and not np.isfinite(padeye_z):
    padeye_z = -raster_depth if np.isfinite(raster_depth) else padeye_z
  anchor_bottom_z = data.get("anchor_bottom_z", np.nan)
  try:
    anchor_bottom_z = float(anchor_bottom_z)
  except (TypeError, ValueError):
    anchor_bottom_z = np.nan

  info = [
      f"Mooring System: {system_type}",
      "-" * 36,
      "GENERAL",
      f" • Profile: {profile_label}" if profile_label else None,
      f" • Turbines: {num_turbines}",
      f" • Lines: {num_lines}",
      f" • Fairlead Depth: {fairlead_depth:.2f} m" if np.isfinite(fairlead_depth) else " • Fairlead Depth: --",
      f" • Raster Depth: {raster_depth:.2f} m" if np.isfinite(raster_depth) else " • Raster Depth: --",
      "-" * 36,
      "MOORING LINE",
      f" • Radius: {anchor_radius:.2f} m",
      f" • Total Line Length: {line_length:.2f} m",
      f" • Minimum Geometric Length: {min_geom:.2f} m" if np.isfinite(min_geom) else " • Minimum Geometric Length: --",
      f" • Excess Line: {excess:.2f} m" if np.isfinite(excess) else " • Excess Line: --",
      f" • Ground Ratio: {ground_ratio*100.0:.1f}%",
      (f" • Suspended Line Angle: {float(data.get('suspended_line_angle_deg')):.1f}° from horizontal"
       if np.isfinite(float(data.get('suspended_line_angle_deg', np.nan))) else " • Suspended Line Angle: --"),
      f" • Subsurface / Embedded: {L_sub:.2f} m",
      f" • Grounded on Seabed: {L_bot:.2f} m",
      f" • Suspended: {L_sus:.2f} m",
  ]
  info = [x for x in info if x is not None]

  _show_ljp = (
      system_type in ("Catenary", "Taut")
      or (system_type == "Semi-Taut" and not bool(data.get("upper_equals_lower", True)))
  ) and bool(data.get("inverted_bridle", False))
  if _show_ljp:
    _ljp = data.get("lower_joint_pos", np.nan)
    try:
      _ljp = float(_ljp)
    except Exception:
      _ljp = np.nan
    if np.isfinite(_ljp):
      info.append(f" • Lower Joint Position: {_ljp*100.0:.1f}% (Fairlead → TDP)")

  if system_type == "Semi-Taut":
    info.extend([
        " • Suspended Breakdown:",
        f"   • Top Chain: {float(data.get('L_top_chain', 0.0)):.2f} m",
        f"   • Middle Rope: {float(data.get('L_taut', 0.0)):.2f} m",
        f"   • Bottom Chain: {float(data.get('L_bottom_chain_suspended', 0.0)):.2f} m",
        f"   • Taut Percentage: {float(data.get('taut_percentage', 0.0)):.1f}%",
    ])

  info.extend(["-" * 36, "ANCHOR", f" • Type: {primary_anc}"])
  if primary_anc in ("DEA", "Gravity"):
    info.extend([f" • Width: {anchor_width:.2f} m", f" • Height: {anchor_height:.2f} m"])
  else:
    # For cylindrical/pile anchors, Excel stores penetration depth and
    # physical anchor length separately.  Report the physical length here;
    # anchor_depth remains the seabed-to-bottom penetration used by geometry.
    info.extend([f" • Diameter: {anchor_width:.2f} m", f" • Length: {anchor_height:.2f} m"])
  if data.get("chain_d"):
    info.append(f" • Chain Diameter: {float(data.get('chain_d'))*1000.0:.1f} mm")
  if data.get("rope_d"):
    info.append(f" • Rope Diameter: {float(data.get('rope_d'))*1000.0:.1f} mm")
  if system_type != "Taut" and padeye_angle is not None:
    info.append(f" • Padeye Angle: {float(padeye_angle):.2f}° from horizontal")
  if padeye_fraction is not None:
    info.append(f" • Padeye Position: {100.0*float(padeye_fraction):.1f}% down from top")
  if np.isfinite(padeye_z):
    info.append(f" • Padeye Elevation: {padeye_z:.2f} m")
  # Report the same maximum-penetration Z used by the 2-D reference line.
  # Taut uses physical anchor length; Catenary/Semi-Taut use the configured
  # penetration depth. Keeping the calculation here identical prevents the
  # infobox and profile annotation from disagreeing.
  pen_depth_for_info = (
      get_taut_effective_penetration_depth(anchor_depth, anchor_height)
      if system_type == "Taut" else max(anchor_depth, 0.0)
  )
  if np.isfinite(anchor_bottom_z):
    max_penetration_z = anchor_bottom_z
  elif np.isfinite(padeye_z):
    max_penetration_z = padeye_z - pen_depth_for_info
  elif np.isfinite(raster_depth):
    max_penetration_z = -raster_depth - pen_depth_for_info
  else:
    max_penetration_z = np.nan
  if np.isfinite(max_penetration_z):
    _max_penetration_label = (
        "Branch Anchor Bottom / Max Penetration"
        if data.get("bridle_selected_branch") else "Max Penetration"
    )
    info.append(f" • {_max_penetration_label}: z = {max_penetration_z:.2f} m")

  T_fairlead = float(data.get("T_fairlead", np.hypot(data.get("HF",0.0), data.get("VF",0.0))))
  T_padeye = float(data.get("T_padeye", np.hypot(data.get("HA",0.0), data.get("VA",0.0))))
  T_max = float(data.get("T_max", max(T_fairlead, T_padeye)))
  if system_type == "Semi-Taut":
    T_max = max(T_max, float(data.get("taut_max_tension", T_max)))
  elif system_type == "Taut":
    T_max = max(T_max, float(data.get("taut_max_tension", T_max)))

  info.extend([
      "-" * 36,
      "TENSIONS AND FOS",
      f" • Fairlead: {T_fairlead/1000.0:.2f} kN (H={float(data.get('HF',0.0))/1000.0:.2f}, V={float(data.get('VF',0.0))/1000.0:.2f})",
      f" • Anchor / Padeye: {T_padeye/1000.0:.2f} kN (H={float(data.get('HA',0.0))/1000.0:.2f}, V={float(data.get('VA',0.0))/1000.0:.2f})",
      f" • Mooring Line T(max): {T_max/1000.0:.2f} kN",
  ])
  if system_type in ("Semi-Taut", "Taut") and np.isfinite(float(data.get("taut_min_tension", np.nan))):
    info.append(f" • Taut Minimum: {float(data.get('taut_min_tension'))/1000.0:.2f} kN")
    info.append(f" • Taut Maximum: {float(data.get('taut_max_tension', T_max))/1000.0:.2f} kN")
  if system_type == "Semi-Taut":
    info.append(f" • Chain FoS: {float(data.get('line_fos_chain', 0.0)):.2f} | Rope FoS: {float(data.get('line_fos_rope', 0.0)):.2f}")
  elif system_type == "Taut":
    info.append(f" • Rope FoS: {float(data.get('line_fos', 0.0)):.2f}")
  else:
    info.append(f" • Line FoS: {float(data.get('line_fos_fairlead', 0.0)):.2f} / {float(data.get('line_fos_padeye', 0.0)):.2f}")

  if data.get("anchor_results"):
    info.append(" • Anchor Suitability:")
    for aname, res in data["anchor_results"].items():
      info.append(f"   - {aname}: [{res.get('status','--')}] FoS: {float(res.get('fos',0.0)):.2f}")
  return info


def _add_inverted_bridle_info_box_2d(ax, data, result):
  """Display inverted-bridle inputs/results on a 2D profile."""
  if not result or not result.get("success"):
    return

  inverted = bool(data.get("inverted_bridle", False))
  triad = bool(data.get("triad", False))

  lines = []
  if inverted:
    lower = result.get(
        "lower_joint_position_fraction",
        data.get("lower_joint_pos")
    )
    angle = result.get(
        "inter_anchor_angle",
        data.get("inter_anchor_angle")
    )
    lines.append("Inverted Bridle: YES")
    if lower is not None:
      lines.append(f"Lower Joint position: {float(lower) * 100.0:.1f}%")
    if angle is not None:
      lines.append(f"Inter-Anchor Angle: {float(angle):.1f}°")
    lines.append("Triad: NO")
  else:
    lines.append("Inverted Bridle: NO")

  ax.text(
      0.02, 0.98,
      "\n".join(lines),
      transform=ax.transAxes,
      va="top",
      ha="left",
      fontsize=9,
      bbox=dict(boxstyle="round,pad=0.4", facecolor="white", alpha=0.85),
      zorder=100,
  )


def _add_inverted_bridle_info_box_3d(plotter, data, result):
  """Display inverted-bridle inputs/results in the PyVista view."""
  if not result or not result.get("success"):
    return
  if not bool(data.get("inverted_bridle", False)):
    return

  lower = result.get(
      "lower_joint_position_fraction",
      data.get("lower_joint_pos")
  )
  angle = result.get(
      "inter_anchor_angle",
      data.get("inter_anchor_angle")
  )

  lines = ["Inverted Bridle: YES"]
  if lower is not None:
    lines.append(f"Lower Joint position: {float(lower) * 100.0:.1f}%")
  if angle is not None:
    lines.append(f"Inter-Anchor Angle: {float(angle):.1f}°")

  try:
    plotter.add_text(
        "\n".join(lines),
        position="upper_left",
        font_size=10,
        name="inverted_bridle_info",
    )
  except Exception:
    pass


def render_axes_content(ax1, ax2, data):
  x_plot = data["x_plot"]
  z_plot = data["z_plot"]
  xf = data["xf"]
  water_depth = data["water_depth"]
  fairlead_draft = data["fairlead_draft"]
  anchor_radius = data["anchor_radius"]
  line_length = data["line_length"]
  num_lines = data["num_lines"]
  num_turbines = data["num_turbines"]
  farm_area = data["farm_area"]
  buffer_zone = data.get("buffer_zone", 50.0)
  L_sus = data["L_sus"]
  L_bot = data["L_bot"]
  HF = data["HF"]
  VF = data["VF"]
  HA = data["HA"]
  VA = data["VA"]
  X_td = data["X_td"]
  center_lat = data["center_lat"]
  center_lon = data["center_lon"]
  sub_x = data.get("sub_x")
  sub_z = data.get("sub_z")
  L_sub = data.get("L_sub", 0.0)
  rope_nodes = data.get("rope_nodes")
  system_type = data["system_type"]
  anchor_results = data["anchor_results"]
  z_anchor_custom = data.get("z_anchor_custom")
  anchor_width = data["anchor_width"]
  anchor_height = data["anchor_height"]
  anchor_depth = data["anchor_depth"]
  primary_anc = data["primary_anc"]

  # ---------------------------------------------------------
  # 2D PROFILE PLOT
  # ---------------------------------------------------------
  ax1.axhline(
      0,
      color="blue",
      linestyle="--",
      linewidth=1.5,
      zorder=1,
      label="Sea Level ($z = 0$ m)",
  )
  ax1.axhline(
      -water_depth,
      color="saddlebrown",
      linestyle="-",
      linewidth=3.5,
      zorder=1,
      label=f"Seabed ($z = {-water_depth}$ m)",
  )

  z_anchor_bottom = -water_depth - anchor_depth
  ax1.axhline(
      z_anchor_bottom,
      color="purple",
      linestyle=":",
      linewidth=1.5,
      zorder=1,
      label="_nolegend_",
  )
  if system_type == "Taut":
    z_anchor = float(selected.get("padeye_z", selected.get("seabed_z", z_anchor_custom if z_anchor_custom is not None else -water_depth)))
    ax1.plot(x_plot, z_plot, color="black", linewidth=6.0, zorder=2)
    ax1.plot(
        x_plot,
        z_plot,
        color="deepskyblue",
        linewidth=4.0,
        zorder=3,
        label=f"Taut Mooring Line ({data.get('rope_material', 'HMPE')})",
    )
    ax1.plot(
        0.0,
        z_anchor,
        "ro",
        markeredgecolor="black",
        markersize=5,
        zorder=5,
        label="Padeye",
    )
  else:
    z_anchor = sub_z[0] if sub_z is not None else -water_depth
    ax1.plot(
        0.0,
        z_anchor,
        "ro",
        markeredgecolor="black",
        markersize=5,
        zorder=5,
        label="Padeye",
    )

    if (
        system_type == "Semi-Taut"
        and rope_nodes is not None
        and not (selected.get("bridle_geometry") and selected.get("bridle_geometry").get("enabled"))
    ):
      plot_semi_taut_2d_sections(ax1, data, data, selected_ground=None, labels=True)
    elif (
        rope_nodes is not None
        and not (selected.get("bridle_geometry") and selected.get("bridle_geometry").get("enabled"))
    ):
      x1_n, z1_n, x2_n, z2_n = rope_nodes
      mask_bot = x_plot <= x1_n
      bot_x = (np.concatenate((sub_x, x_plot[mask_bot][1:] if np.any(mask_bot) else []))
               if sub_x is not None else x_plot[mask_bot])
      bot_z = (np.concatenate((sub_z, z_plot[mask_bot][1:] if np.any(mask_bot) else []))
               if sub_z is not None else z_plot[mask_bot])
      ax1.plot(bot_x, bot_z, "k-", linewidth=2.5, zorder=2, label="Mooring Line Profile (Chain)")
      ax1.plot([x1_n, x2_n], [z1_n, z2_n], color="black", linewidth=6.0, zorder=2)
      ax1.plot([x1_n, x2_n], [z1_n, z2_n], color="deepskyblue", linewidth=4.0, zorder=3,
               label=f"Synthetic Section ({data.get('rope_material', 'Polyester')})")
      mask_top = x_plot >= x2_n
      if np.any(mask_top):
        ax1.plot(np.concatenate(([x2_n], x_plot[mask_top], [xf])),
                 np.concatenate(([z2_n], z_plot[mask_top], [fairlead_draft])),
                 "k-", linewidth=2.5, zorder=2)
      ax1.plot(x1_n, z1_n, "D", color="magenta", markeredgecolor="black", markersize=5, zorder=50, label="Bottom Joint")
      ax1.plot(x2_n, z2_n, "s", color="magenta", markeredgecolor="black", markersize=5, zorder=50, label="Top Joint")
    else:
      _legacy_bridle = selected.get("bridle_geometry")
      if (
          _legacy_bridle
          and _legacy_bridle.get("enabled")
          and system_type == "Semi-Taut"
          and rope_nodes is not None
      ):
        x1_n, z1_n, x2_n, z2_n = map(float, rope_nodes)
        rx, rz = _section_curve(x_plot, z_plot, x1_n, x2_n)
        if rx.size:
          ax1.plot(rx, rz, color="black", linewidth=6.0, zorder=2)
          ax1.plot(rx, rz, color="deepskyblue", linewidth=4.0, zorder=3,
                   label=f"Synthetic Rope ({data.get('rope_material', 'Polyester')})")
        tx, tz = _section_curve(x_plot, z_plot, x2_n, xf)
        if tx.size:
          tx[-1] = xf
          tz[-1] = fairlead_draft
          ax1.plot(tx, tz, "k-", linewidth=2.6, zorder=2, label="Top Chain")
        ax1.plot(x1_n, z1_n, "D", color="magenta", markeredgecolor="black",
                 markersize=6, zorder=50, label="Lower Joint")
        ax1.plot(x2_n, z2_n, "s", color="magenta", markeredgecolor="black",
                 markersize=6, zorder=50, label="Upper Joint")
      else:
        if sub_x is not None and sub_z is not None:
          full_x = np.concatenate((sub_x, x_plot[1:]))
          full_z = np.concatenate((sub_z, z_plot[1:]))
        else:
          full_x, full_z = x_plot, z_plot
        ax1.plot(full_x, full_z, "k-", linewidth=2.5, zorder=2, label="Mooring Line Profile")

    if sub_x is not None and sub_z is not None:
      ddp_x = sub_x[-1]
      ddp_z = sub_z[-1]
      ax1.plot(
          ddp_x,
          ddp_z,
          "s",
          color="yellow",
          markeredgecolor="black",
          markersize=6,
          zorder=6,
          label="Down Dip Point (DDP)",
      )

  # Inverted Bridle 2-D profile: reuse the EXISTING solved lower profile.
  # H25 is a plan-view angle, so it must not create a new vertical/profile
  # angle. In section both branches therefore retain the original chain/
  # catenary elevation curve and meet at the same Joint 1.
  bridle = selected.get("bridle_geometry")
  if bridle and bridle.get("enabled"):
    bx, bz = bridle["joint_local"]
    branch_profiles = bridle.get("profiles", [])
    for k, branch in enumerate(branch_profiles, start=1):
      ax1.plot(
          branch["profile_x"], branch["profile_z"],
          "k-", linewidth=2.5, zorder=4,
          label=(
              f"T{selected_turbine_index + 1}L{selected_line_index + 1}{'a' if k == 1 else 'b'}"
              if (not bridle_navigation or k == selected_branch_index) else "_nolegend_"
          ),
      )
    # Semi-Taut already renders the physical Bottom Chain / Rope Joint above.
    # Do not add a second generic "Lower Joint" symbol.
    if system_type != "Semi-Taut":
      ax1.plot(bx, bz, "D", color="magenta", markeredgecolor="black",
               markersize=6, zorder=50, label="Lower Joint")

  # Taut is always center-aligned, others are inside-aligned
  is_top_mid = system_type == "Taut"
  rect_x = -anchor_width / 2.0 if is_top_mid else -anchor_width
  rect_z = (
      z_anchor - anchor_height
      if is_top_mid
      else z_anchor - (anchor_height / 2.0)
  )

  rect_2d = patches.Rectangle(
      (rect_x, rect_z),
      anchor_width,
      anchor_height,
      linewidth=1.5,
      edgecolor="black",
      facecolor="gray",
      zorder=4,
      label="Anchor",
  )
  ax1.add_patch(rect_2d)
  ax1.plot(
      xf,
      fairlead_draft,
      "go",
      markeredgecolor="black",
      markersize=5,
      zorder=4,
      label="Fairlead/Turbine",
  )

  if X_td is not None and L_bot > 0:
    ax1.plot(
        X_td,
        -water_depth,
        "o",
        color="yellow",
        markeredgecolor="black",
        markersize=6,
        zorder=4,
        label="Touchdown Point (TDP)",
    )

  info = build_mooring_info_lines(
      data, profile_label=data.get("profile_label")
  )
  ax1.text(
      0.03,
      0.97,
      "\n".join(info),
      transform=ax1.transAxes,
      fontsize=6.5,
      va="top",
      bbox=dict(boxstyle="round", facecolor="white", alpha=0.9),
  )
  ax1.set_title(f"Mooring Line Profile ({primary_anc})", fontweight="bold")
  ax1.set_xlabel("Horizontal Distance [m]")
  ax1.set_ylabel("Elevation [m]")
  ax1.set_xlim(-anchor_width - 15, xf + 25)
  ax1.grid(True, linestyle=":", alpha=0.6)
  _dedupe_mooring_legend(ax1)

  # ---------------------------------------------------------
  # 3D FARM LAYOUT IN REAL CRS EPSG:3857
  # ---------------------------------------------------------
  origin_x, origin_y = latlon_to_epsg3857(center_lat, center_lon)
  min_spacing = (2.0 * anchor_radius) + buffer_zone
  base_spacing = (
      np.sqrt((farm_area * 1e6) / num_turbines)
      if farm_area and farm_area > 0
      else min_spacing
  )
  spacing = max(min_spacing, base_spacing)

  rows, cols = get_optimal_grid_dimensions(num_turbines)
  total_width_x = cols * spacing
  total_height_y = rows * spacing * 0.866

  headings = generate_mooring_headings(
      num_lines,
      triad=bool(data.get("triad", False)),
      inter_arm_angle_deg=float(data.get("inter_arm_angle", 0.0)),
      angle_between_mooring_arms_deg=float(
          data.get("angle_between_mooring_arms", 360.0 / max(num_lines, 1))
      ),
  )
  base_circle_angles = np.linspace(0, 2 * np.pi, 360)
  circle_angles = np.sort(
      np.unique(np.concatenate((base_circle_angles, np.radians(headings))))
  )

  # Align 3D space with the new padded logic
  r_padeye = (
      anchor_radius - (anchor_width / 2.0)
      if system_type == "Taut"
      else anchor_radius - anchor_width
  )
  r_fairlead = r_padeye - xf
  r_anchor = anchor_radius

  r_tdp = r_padeye - X_td if X_td is not None else None
  r_ddp = (
      r_padeye - sub_x[-1] if (sub_x is not None and len(sub_x) > 0) else None
  )

  if system_type == "Taut":
    # Taut padeye/head sits at the seabed surface; anchor bottom defines max penetration.
    z_anchor_3d = (
        z_anchor_custom
        if z_anchor_custom is not None
        else -water_depth
    )
  else:
    z_anchor_3d = (
        sub_z[0]
        if (sub_z is not None and len(sub_z) > 0)
        else (-water_depth - anchor_depth + 0.5 * anchor_height)
    )

  turbine_coords = []
  for i in range(num_turbines):
    row_idx = i // cols
    col_idx = i % cols
    cx = origin_x - (total_width_x / 2.0) + (col_idx + 0.5 * (row_idx % 2)) * spacing
    cy = origin_y - (total_height_y / 2.0) + row_idx * spacing * 0.866
    turbine_coords.append((cx, cy))

    ax2.scatter(
        cx,
        cy,
        fairlead_draft,
        color="blue",
        s=40,
        marker="^",
        edgecolor="black",
        linewidths=0.8,
        zorder=5,
        label=("Floating Turbine" if i == 0 else ""),
    )

    ax2.plot(
        cx + r_anchor * np.cos(circle_angles),
        cy + r_anchor * np.sin(circle_angles),
        np.full_like(circle_angles, -water_depth),
        "r--",
        lw=0.8,
        alpha=0.6,
        label="Anchor Radius" if i == 0 else "",
    )

    if r_tdp is not None and r_tdp > 0:
      ax2.plot(
          cx + r_tdp * np.cos(circle_angles),
          cy + r_tdp * np.sin(circle_angles),
          np.full_like(circle_angles, -water_depth),
          color="gold",
          linestyle="--",
          lw=1.0,
          alpha=0.8,
          label="TDP Radius" if i == 0 else "",
      )

    if r_ddp is not None and r_ddp > 0:
      ax2.plot(
          cx + r_ddp * np.cos(circle_angles),
          cy + r_ddp * np.sin(circle_angles),
          np.full_like(
              circle_angles, sub_z[-1] if sub_z is not None else -water_depth
          ),
          color="orange",
          linestyle="--",
          lw=1.0,
          alpha=0.8,
          label="DDP Radius" if i == 0 else "",
      )

    for angle in headings:
      angle_rad = np.radians(angle)
      ux, uy = np.cos(angle_rad), np.sin(angle_rad)
      fl_x, fl_y = cx + r_fairlead * ux, cy + r_fairlead * uy
      ax2.plot(
          [cx, fl_x],
          [cy, fl_y],
          [fairlead_draft, fairlead_draft],
          color="gray",
          lw=0.8,
          alpha=0.5,
      )

      # FIX: Ensure 3D Padeye is plotted exactly at r_padeye instead of r_anchor
      padeye_3d_x = cx + r_padeye * ux
      padeye_3d_y = cy + r_padeye * uy
      ax2.scatter(
          [padeye_3d_x],
          [padeye_3d_y],
          [z_anchor_3d],
          color="red",
          marker="o",
          s=18,
          edgecolors="black",
          linewidths=0.5,
          zorder=6,
          label="Padeye" if (i == 0 and angle == headings[0]) else "",
      )

      # =========================================================================
      # NEW 3D ANCHOR VISUALIZATION GENERATOR
      # =========================================================================
      anc_type_upper = primary_anc.upper()
      is_taut = system_type == "Taut"

      # Determine explicit geometries and dimensions based on anchor type
      if any(
          k in anc_type_upper for k in ["DRIVEN", "DRILLED", "SUCTION", "PILE"]
      ):
        is_cylinder = True
        cyl_dia = anchor_width
        cyl_len = anchor_height
      else:
        is_cylinder = False
        box_w = anchor_width
        box_l = anchor_width
        box_h = anchor_height

      # Calculate Anchor Center Coordinates (cx_anc, cy_anc, z_top, z_bot) based on System constraints
      if is_taut:
        # Taut: Padeye acts like a hat on top center
        cx_anc = padeye_3d_x
        cy_anc = padeye_3d_y
        if is_cylinder:
          z_top = z_anchor_3d
          z_bot = z_anchor_3d - cyl_len
        else:
          z_top = z_anchor_3d
          z_bot = z_anchor_3d - box_h
      else:
        # Catenary/Semi-Taut: keep the cylindrical anchor body fixed while
        # placing the padeye at the Excel-specified fraction down from the top.
        if is_cylinder and any(k in anc_type_upper for k in ("SUCTION", "DRIVEN", "DRILLED")):
          frac = float((data.get("padeye_params") or {}).get(primary_anc, {}).get("position_fraction", 0.50))
          z_bot, z_top, _ = get_pile_anchor_vertical_geometry(
              water_depth, anchor_depth, anchor_height, padeye_fraction=frac
          )
          cx_anc = padeye_3d_x + (cyl_dia / 2.0) * ux
          cy_anc = padeye_3d_y + (cyl_dia / 2.0) * uy
        elif is_cylinder:
          cx_anc = padeye_3d_x + (cyl_dia / 2.0) * ux
          cy_anc = padeye_3d_y + (cyl_dia / 2.0) * uy
          z_top = z_anchor_3d + (cyl_len / 2.0)
          z_bot = z_anchor_3d - (cyl_len / 2.0)
        else:
          cx_anc = padeye_3d_x + (box_l / 2.0) * ux
          cy_anc = padeye_3d_y + (box_l / 2.0) * uy
          z_top = z_anchor_3d + (box_h / 2.0)
          z_bot = z_anchor_3d - (box_h / 2.0)

      # Draw shapes using purely wireframes (ax2.plot) so they remain explicitly extractable
      if is_cylinder:
        # Generate Cylinder Wireframe Array
        theta = np.linspace(0, 2 * np.pi, 24)
        Xc = cx_anc + (cyl_dia / 2.0) * np.cos(theta)
        Yc = cy_anc + (cyl_dia / 2.0) * np.sin(theta)
        ax2.plot(
            Xc,
            Yc,
            np.full_like(Xc, z_bot),
            color="purple",
            lw=1.2,
            alpha=0.8,
            label=(
                f"Anchor Body ({primary_anc})"
                if (i == 0 and angle == headings[0])
                else ""
            ),
        )
        ax2.plot(
            Xc, Yc, np.full_like(Xc, z_top), color="purple", lw=1.2, alpha=0.8
        )
        # Vertical struts
        for idx_strut in range(0, len(theta), 6):
          ax2.plot(
              [Xc[idx_strut], Xc[idx_strut]],
              [Yc[idx_strut], Yc[idx_strut]],
              [z_bot, z_top],
              color="purple",
              lw=1.2,
              alpha=0.8,
          )
      else:
        # Generate Cuboid Wireframe Array
        vx, vy = -uy, ux  # Perpendicular vector calculated to form box
        dx = box_l / 2.0
        dy = box_w / 2.0

        # Compute the 5 connecting coordinates for bottom and top rectangle paths
        bx = [
            cx_anc + dx * ux + dy * vx,
            cx_anc - dx * ux + dy * vx,
            cx_anc - dx * ux - dy * vx,
            cx_anc + dx * ux - dy * vx,
            cx_anc + dx * ux + dy * vx,
        ]
        by = [
            cy_anc + dx * uy + dy * vy,
            cy_anc - dx * uy + dy * vy,
            cy_anc - dx * uy - dy * vy,
            cy_anc + dx * uy - dy * vy,
            cy_anc + dx * uy + dy * vy,
        ]

        ax2.plot(
            bx,
            by,
            np.full_like(bx, z_bot),
            color="purple",
            lw=1.2,
            alpha=0.8,
            label=(
                f"Anchor Body ({primary_anc})"
                if (i == 0 and angle == headings[0])
                else ""
            ),
        )
        ax2.plot(
            bx, by, np.full_like(bx, z_top), color="purple", lw=1.2, alpha=0.8
        )
        # Connect top and bottom faces at corners
        for cx_corner, cy_corner in zip(bx[:-1], by[:-1]):
          ax2.plot(
              [cx_corner, cx_corner],
              [cy_corner, cy_corner],
              [z_bot, z_top],
              color="purple",
              lw=1.2,
              alpha=0.8,
          )
      # =========================================================================

      # --- Continue Generating Standard Profiles ---
      if rope_nodes is not None:
        x1_n, z1_n, x2_n, z2_n = rope_nodes
        mask_bot = x_plot <= x1_n
        bot_x = (
            np.concatenate(
                (sub_x, x_plot[mask_bot][1:] if np.any(mask_bot) else [])
            )
            if sub_x is not None
            else x_plot[mask_bot]
        )
        bot_z = (
            np.concatenate(
                (sub_z, z_plot[mask_bot][1:] if np.any(mask_bot) else [])
            )
            if sub_z is not None
            else z_plot[mask_bot]
        )
        if len(bot_x) == 0 or bot_x[-1] != x1_n:
          bot_x = np.append(bot_x, x1_n)
          bot_z = np.append(bot_z, z2_n if len(bot_z) == 0 else bot_z[-1])

        if not selected.get("bridle_geometry"):
          ax2.plot(
              cx + (r_padeye - bot_x[::2]) * ux,
              cy + (r_padeye - bot_x[::2]) * uy,
              bot_z[::2],
              color="black",
              alpha=0.7,
              lw=1.2,
          )
        ax2.plot(
            cx + (r_padeye - np.array([x1_n, x2_n])) * ux,
            cy + (r_padeye - np.array([x1_n, x2_n])) * uy,
            [z1_n, z2_n],
            color="deepskyblue",
            alpha=0.9,
            lw=2.0,
        )
      else:
        full_x = (
            np.concatenate((sub_x, x_plot[1:])) if sub_x is not None else x_plot
        )
        full_z = (
            np.concatenate((sub_z, z_plot[1:])) if sub_z is not None else z_plot
        )
        if selected.get("bridle_geometry"):
          _jx = float(selected["bridle_geometry"]["joint_local"][0])
          _m = full_x <= _jx + 1e-8
          _fx3 = np.concatenate((full_x[_m], [_jx]))
          _fz3 = np.concatenate((full_z[_m], [float(selected["bridle_geometry"]["joint_local"][1])]))
        else:
          _fx3, _fz3 = full_x, full_z
        ax2.plot(
            cx + (r_padeye - _fx3[::10]) * ux,
            cy + (r_padeye - _fx3[::10]) * uy,
            _fz3[::10],
            color="deepskyblue" if "Taut" in system_type else "black",
            alpha=0.5,
            lw=1.2,
        )

      # Inverted Bridle: reuse the original solved lower profile twice.
      # The only change is PLAN position: the two copies run from Joint 1 to
      # the two anchors at +/- H25/2. No straight 3-D connector is generated.
      bridle = local_bridle
      if bridle and bridle.get("enabled") and len(bridle.get("profiles", [])) == 2:
        jx, jy, jz = bridle["joint"]

        for b_idx, branch in enumerate(bridle["profiles"], start=1):
          _suffix = "a" if b_idx == 1 else "b"
          _line_label = f"T{i+1}L{j+1}{_suffix}"

          _bx = np.asarray(branch.get("x", []), dtype=float)
          _by = np.asarray(branch.get("y", []), dtype=float)
          _bz = np.asarray(branch.get("z", []), dtype=float)

          if _bx.size >= 2 and _by.size == _bx.size and _bz.size == _bx.size:
            ax2.plot(
                _bx,
                _by,
                _bz,
                color="deepskyblue" if system_type == "Taut" else "black",
                alpha=0.98,
                lw=2.8,
                zorder=9,
                label=_line_label
                if (i == 0 and angle == headings[0]) else "",
            )

          if i == 0 and angle == headings[0]:
            a = bridle["anchors"][b_idx - 1]
            ax2.scatter(
                [a["x"]], [a["y"]], [a["z"]],
                color="red",
                marker="o",
                s=24,
                edgecolors="black",
                linewidths=0.7,
                zorder=10,
                label=f"{_line_label} Anchor",
            )

        # Semi-Taut already has the physical Bottom Chain / Rope Joint marker.
        if system_type != "Semi-Taut":
          ax2.scatter(
              [jx], [jy], [jz],
              color="magenta",
              marker="D",
              s=30,
              edgecolors="black",
              linewidths=0.7,
              zorder=10,
              label="Lower Joint"
              if (i == 0 and angle == headings[0]) else "",
          )

      if X_td is not None and r_tdp is not None:
        td_3d_x = cx + r_tdp * ux
        td_3d_y = cy + r_tdp * uy
        ax2.scatter(
            [td_3d_x],
            [td_3d_y],
            [-water_depth],
            color="yellow",
            s=18,
            edgecolors="black",
            linewidths=0.5,
            zorder=6,
            label=(
                "Touchdown Point (TDP)"
                if (i == 0 and angle == headings[0])
                else ""
            ),
        )

      if sub_x is not None and sub_z is not None and r_ddp is not None:
        ddp_3d_x = cx + r_ddp * ux
        ddp_3d_y = cy + r_ddp * uy
        ddp_3d_z = sub_z[-1]
        ax2.scatter(
            [ddp_3d_x],
            [ddp_3d_y],
            [ddp_3d_z],
            color="yellow",
            marker="s",
            s=18,
            edgecolors="black",
            linewidths=0.5,
            zorder=6,
            label=(
                "Down Dip Point (DDP)"
                if (i == 0 and angle == headings[0])
                else ""
            ),
        )

  if turbine_coords:
    min_cx = min(tc[0] for tc in turbine_coords)
    max_cx = max(tc[0] for tc in turbine_coords)
    min_cy = min(tc[1] for tc in turbine_coords)
    max_cy = max(tc[1] for tc in turbine_coords)
    buffer = anchor_radius
    xmin, xmax, ymin, ymax = (
        min_cx - buffer,
        max_cx + buffer,
        min_cy - buffer,
        max_cy + buffer,
    )
    num_p = 30
    edge_x = np.concatenate([
        np.linspace(xmin, xmax, num_p),
        np.full(num_p, xmax),
        np.linspace(xmax, xmin, num_p),
        np.full(num_p, xmin),
    ])
    edge_y = np.concatenate([
        np.full(num_p, ymin),
        np.linspace(ymin, ymax, num_p),
        np.full(num_p, ymax),
        np.linspace(ymax, ymin, num_p),
    ])
    ax2.plot(
        edge_x,
        edge_y,
        np.full_like(edge_x, -water_depth),
        color="purple",
        linestyle="-",
        linewidth=1.5,
        alpha=0.7,
        label="Farm Perimeter",
    )

  span = max(total_width_x, total_height_y) + (2.0 * anchor_radius) + 5000.0
  ax2.set_xlim(origin_x - span / 2, origin_x + span / 2)
  ax2.set_ylim(origin_y - span / 2, origin_y + span / 2)
  ax2.set_zlim(-water_depth - 35, 50.0)
  ax2.set_box_aspect([1, 1, 0.5])
  ax2.set_title(f"Farm Layout ({primary_anc})", fontweight="bold")
  _dedupe_mooring_legend(ax2)




  try:
    _add_inverted_bridle_info_box_2d(ax1, data, data)
  except Exception:
    pass

def read_common_excel_params(sheet):
  # Cell locations are controlled exclusively by EXCEL_CELLS above.
  # Keep the existing return order/meaning unchanged.
  water_depth = excel_cell(sheet, "water_depth", default=120.0, cast_type=float)
  num_turbines = excel_cell(sheet, "number_of_turbines", default=100, cast_type=int)
  farm_area = excel_cell(sheet, "farm_area", default=429.18, cast_type=float)
  buffer_zone = excel_cell(sheet, "buffer_zone", default=50.0, cast_type=float)
  center_lat = excel_cell(sheet, "center_latitude", default=57.1497, cast_type=float)
  center_lon = excel_cell(sheet, "center_longitude", default=-2.0943, cast_type=float)
  print(
      f"[INFO] Read common params -> Depth: {water_depth}, Turbines:"
      f" {num_turbines}, Area: {farm_area}, Buffer: {buffer_zone}, Lat:"
      f" {center_lat}, Lon: {center_lon}"
  )
  return (
      water_depth,
      num_turbines,
      farm_area,
      buffer_zone,
      center_lat,
      center_lon,
  )



def main(excel_path=None):
  """Launch the dashboard workflow using a user-selected Excel database."""
  if excel_path is None:
    excel_path = launch_excel_welcome_screen(
        initial_excel_path=_SELECTED_EXCEL_WORKBOOK_PATH or ""
    )
    if not excel_path:
      return

  excel_path = _normalise_excel_workbook_path(excel_path)
  _set_selected_excel_workbook_path(excel_path)

  print(
      "========================================================================="
  )
  print(" Welcome to Nurdins Mooring Solution (Excel-Driven Multi-Simulation)")
  print(
      "========================================================================="
  )

  try:
    sheet = get_excel_sheet("Mooring_Python", excel_path=excel_path)
  except Exception as e:
    print(f"[ERROR] Could not connect to Excel sheet 'Mooring_Python': {e}")
    return

  # Reuse the already-connected worksheet and its shared parameters for all
  # three systems.  This avoids repeated Excel/COM lookups during startup.
  try:
    common_excel_params = read_common_excel_params(sheet)
  except Exception as e:
    print(f"[ERROR] Could not read common Excel parameters: {e}")
    return

  anchor_db, all_anchors = read_anchor_database(sheet)
  try:
    padeye_params = read_padeye_parameters(sheet)
  except Exception as e:
    print(f"[ERROR] Invalid padeye parameters from Excel: {e}")
    return
  raster_depth_threshold = read_raster_depth_threshold(sheet)
  _threshold_fraction = normalise_raster_depth_threshold(raster_depth_threshold)
  print(
      f"[INFO] Raster depth suitability threshold from H30: "
      f"+/-{float(_threshold_fraction) * 100.0:.2f}%"
  )

  # Read system execution flags from Excel
  cat_flag = bool(excel_cell(sheet, "run_catenary", default=1, cast_type=int))
  semi_flag = bool(excel_cell(sheet, "run_semi_taut", default=1, cast_type=int))
  taut_flag = bool(excel_cell(sheet, "run_taut", default=1, cast_type=int))

  # User-defined mooring-line azimuth arrangement.
  triad = bool(excel_cell(sheet, "triad", default=0, cast_type=int))
  inter_arm_angle = excel_cell(
      sheet, "inter_arm_angle", default=10.0, cast_type=float
  )
  angle_between_mooring_arms = excel_cell(
      sheet, "angle_between_mooring_arms", default=120.0, cast_type=float
  )

  upper_equals_lower = bool(excel_cell(sheet, "upper_equals_lower", default=1, cast_type=int))
  lower_joint_pos = excel_cell(sheet, "lower_joint_pos", default=0.50, cast_type=float)
  inverted_bridle = bool(excel_cell(sheet, "inverted_bridle", default=0, cast_type=int))
  inter_anchor_angle = excel_cell(sheet, "inter_anchor_angle", default=30.0, cast_type=float)

  validate_inverted_bridle_configuration(
      triad, upper_equals_lower, inverted_bridle, inter_anchor_angle, lower_joint_pos
  )
  if inverted_bridle:
    print(
        f"[INFO] Inverted Bridle ENABLED -> H25={inter_anchor_angle:.3f}°, "
        f"H19 Upper=Lower={upper_equals_lower}, H21={lower_joint_pos:.3f}"
    )

  print(
      f"[INFO] Mooring angle configuration -> Triad={triad}, "
      f"Inter-arm angle={inter_arm_angle:.3f}°, "
      f"Angle between mooring arms={angle_between_mooring_arms:.3f}°"
  )

  system_flags = {
      "Catenary": cat_flag,
      "Semi-Taut": semi_flag,
      "Taut": taut_flag,
  }

  results_store = {"Catenary": {}, "Semi-Taut": {}, "Taut": {}}

  # -------------------------------------------------------------------------
  # 1. EVALUATE CATENARY SYSTEM
  # -------------------------------------------------------------------------
  if system_flags["Catenary"]:
    print("\n[INFO] Running Catenary Mooring Simulations...")
    cat_params = read_from_file_catenary(sheet, common_excel_params)
    if cat_params:
      (
          water_depth,
          anchor_radius,
          line_length,
          chain_d,
          fairlead_radius,
          fairlead_draft,
          num_turbines,
          farm_area,
          num_lines,
          center_lat,
          center_lon,
          buffer_zone,
      ) = cat_params

      validate_mooring_angle_configuration(
          num_lines, triad, inter_arm_angle, angle_between_mooring_arms
      )

      for anc in all_anchors:
        res = evaluate_catenary_anchor(
            water_depth,
            anchor_radius,
            line_length,
            chain_d,
            fairlead_radius,
            fairlead_draft,
            num_turbines,
            farm_area,
            num_lines,
            center_lat,
            center_lon,
            buffer_zone,
            anchor_db,
            anc,
            padeye_params=padeye_params,
        )
        res["padeye_params"] = padeye_params
        res.setdefault("water_depth", water_depth)
        res.setdefault("raster_depth_threshold", raster_depth_threshold)
        res.setdefault("anchor_radius", anchor_radius)
        res.setdefault("fairlead_radius", fairlead_radius)
        res.setdefault("fairlead_draft", fairlead_draft)
        res.setdefault("line_length", line_length)
        res.setdefault("chain_d", chain_d)
        res.setdefault("system_type", "Catenary")
        res.setdefault("triad", triad)
        res.setdefault("inter_arm_angle", inter_arm_angle)
        res.setdefault("angle_between_mooring_arms", angle_between_mooring_arms)
        res.setdefault("upper_equals_lower", upper_equals_lower)
        res.setdefault("lower_joint_pos", lower_joint_pos)
        res.setdefault("inverted_bridle", inverted_bridle)
        res.setdefault("inter_anchor_angle", inter_anchor_angle)
        res.setdefault("primary_anc", anc)
        if inverted_bridle and res.get("success"):
          try:
            _bg = calculate_inverted_bridle_geometry(res, 0.0, (0.0, 0.0))
            res["bridle_geometry_template"] = _bg
            res["bridle_loads"] = calculate_inverted_bridle_loads(res, _bg)
            res["bridle_anchor_results"] = {}
            for _bi, _br in enumerate((res["bridle_loads"] or {}).get("branches", []), start=1):
              res["bridle_anchor_results"][f"Branch {_bi}"] = evaluate_anchor_capacities(anchor_db, [anc], _br["HA"], _br["VA"]).get(anc, {})
          except Exception as _be:
            res["bridle_geometry_error"] = str(_be)
        res["_failure_stage"] = infer_failure_stage("Catenary", res)
        results_store["Catenary"][anc] = res
        status = (
            "PASS" if res.get("success") else f"FAIL ({res.get('msg', '')})"
        )
        print(f"  • Catenary + {anc}: {status}")

  # -------------------------------------------------------------------------
  # 2. EVALUATE SEMI-TAUT SYSTEM
  # -------------------------------------------------------------------------
  if system_flags["Semi-Taut"]:
    print("\n[INFO] Running Semi-Taut Mooring Simulations...")
    semi_params = read_from_file_semi_taut(sheet, common_excel_params)
    if semi_params:
      (
          water_depth,
          anchor_radius,
          line_length,
          rope_d,
          chain_d,
          fairlead_radius,
          fairlead_draft,
          num_turbines,
          farm_area,
          num_lines,
          center_lat,
          center_lon,
          buffer_zone,
          taut_percentage,
      ) = semi_params

      validate_mooring_angle_configuration(
          num_lines, triad, inter_arm_angle, angle_between_mooring_arms
      )

      for anc in all_anchors:
        res = evaluate_semi_taut_anchor(
            water_depth,
            anchor_radius,
            line_length,
            rope_d,
            chain_d,
            fairlead_radius,
            fairlead_draft,
            num_turbines,
            farm_area,
            num_lines,
            center_lat,
            center_lon,
            buffer_zone,
            anchor_db,
            anc,
            padeye_params=padeye_params,
            taut_percentage=taut_percentage,
            upper_equals_lower=upper_equals_lower,
            lower_joint_pos=lower_joint_pos,
        )
        res["padeye_params"] = padeye_params
        res.setdefault("water_depth", water_depth)
        res.setdefault("raster_depth_threshold", raster_depth_threshold)
        res.setdefault("anchor_radius", anchor_radius)
        res.setdefault("fairlead_radius", fairlead_radius)
        res.setdefault("fairlead_draft", fairlead_draft)
        res.setdefault("line_length", line_length)
        res.setdefault("rope_d", rope_d)
        res.setdefault("chain_d", chain_d)
        res.setdefault("taut_percentage", taut_percentage)
        res.setdefault("system_type", "Semi-Taut")
        res.setdefault("triad", triad)
        res.setdefault("inter_arm_angle", inter_arm_angle)
        res.setdefault("angle_between_mooring_arms", angle_between_mooring_arms)
        res.setdefault("upper_equals_lower", upper_equals_lower)
        res.setdefault("lower_joint_pos", lower_joint_pos)
        res.setdefault("inverted_bridle", inverted_bridle)
        res.setdefault("inter_anchor_angle", inter_anchor_angle)
        res.setdefault("primary_anc", anc)
        if inverted_bridle and res.get("success"):
          try:
            _bg = calculate_inverted_bridle_geometry(res, 0.0, (0.0, 0.0))
            res["bridle_geometry_template"] = _bg
            res["bridle_loads"] = calculate_inverted_bridle_loads(res, _bg)
            res["bridle_anchor_results"] = {}
            for _bi, _br in enumerate((res["bridle_loads"] or {}).get("branches", []), start=1):
              res["bridle_anchor_results"][f"Branch {_bi}"] = evaluate_anchor_capacities(anchor_db, [anc], _br["HA"], _br["VA"]).get(anc, {})
          except Exception as _be:
            res["bridle_geometry_error"] = str(_be)
        res["_failure_stage"] = infer_failure_stage("Semi-Taut", res)
        results_store["Semi-Taut"][anc] = res
        status = (
            "PASS" if res.get("success") else f"FAIL ({res.get('msg', '')})"
        )
        print(f"  • Semi-Taut + {anc}: {status}")

  # -------------------------------------------------------------------------
  # 3. EVALUATE TAUT SYSTEM
  # -------------------------------------------------------------------------
  if system_flags["Taut"]:
    print("\n[INFO] Running Taut Mooring Simulations...")
    taut_params = read_from_file_taut(sheet, common_excel_params)
    if taut_params:
      (
          water_depth,
          anchor_radius,
          line_d,
          fairlead_radius,
          fairlead_draft,
          num_turbines,
          farm_area,
          num_lines,
          center_lat,
          center_lon,
          buffer_zone,
      ) = taut_params

      validate_mooring_angle_configuration(
          num_lines, triad, inter_arm_angle, angle_between_mooring_arms
      )

      for anc in all_anchors:
        res = evaluate_taut_anchor(
            water_depth,
            anchor_radius,
            line_d,
            fairlead_radius,
            fairlead_draft,
            num_turbines,
            farm_area,
            num_lines,
            center_lat,
            center_lon,
            buffer_zone,
            anchor_db,
            anc,
        )
        res.setdefault("water_depth", water_depth)
        res.setdefault("raster_depth_threshold", raster_depth_threshold)
        res.setdefault("anchor_radius", anchor_radius)
        res.setdefault("fairlead_radius", fairlead_radius)
        res.setdefault("fairlead_draft", fairlead_draft)
        res.setdefault("rope_d", line_d)
        res.setdefault("system_type", "Taut")
        res.setdefault("triad", triad)
        res.setdefault("inter_arm_angle", inter_arm_angle)
        res.setdefault("angle_between_mooring_arms", angle_between_mooring_arms)
        res.setdefault("upper_equals_lower", upper_equals_lower)
        res.setdefault("lower_joint_pos", lower_joint_pos)
        res.setdefault("inverted_bridle", inverted_bridle)
        res.setdefault("inter_anchor_angle", inter_anchor_angle)
        res.setdefault("rope_material", "HMPE")
        res.setdefault("primary_anc", anc)
        if inverted_bridle and res.get("success"):
          try:
            _bg = calculate_inverted_bridle_geometry(res, 0.0, (0.0, 0.0))
            res["bridle_geometry_template"] = _bg
            res["bridle_loads"] = calculate_inverted_bridle_loads(res, _bg)
            res["bridle_anchor_results"] = {}
            for _bi, _br in enumerate((res["bridle_loads"] or {}).get("branches", []), start=1):
              res["bridle_anchor_results"][f"Branch {_bi}"] = evaluate_anchor_capacities(anchor_db, [anc], _br["HA"], _br["VA"]).get(anc, {})
          except Exception as _be:
            res["bridle_geometry_error"] = str(_be)
        res["_failure_stage"] = infer_failure_stage("Taut", res)
        results_store["Taut"][anc] = res
        status = (
            "PASS" if res.get("success") else f"FAIL ({res.get('msg', '')})"
        )
        print(f"  • Taut + {anc}: {status}")

  # -------------------------------------------------------------------------
  # LAUNCH DASHBOARD
  # -------------------------------------------------------------------------
  print("\n[INFO] Launching Project Setup screen...")
  launch_project_setup_screen(
      results_store,
      system_flags,
      anchor_db,
      water_depth=common_excel_params[0],
      depth_threshold=raster_depth_threshold,
  )


if __name__ == "__main__":
  if len(sys.argv) == 3 and sys.argv[1] == "--seismic-volume-payload":
    run_seismic_volume_payload(sys.argv[2])
  elif len(sys.argv) == 3 and sys.argv[1] == "--seismic-metrics-payload":
    run_seismic_metrics_payload(sys.argv[2])
  else:
    main()

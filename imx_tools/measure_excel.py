from imxInsights import ImxSingleFile
from shapelyM import MeasureLineString
from shapely import Point
import pandas as pd


imx = ImxSingleFile("20250102201919-job-612309.xml")

out_list = []
for imx_object in imx.situation.get_all():
    if isinstance(imx_object.geometry, Point):
        for ref in imx_object.refs:
            ref_field = ref.field
            if ref_field.endswith('@railConnectionRef'):
                rail_con = ref.imx_object
                try:
                    at_measure = imx_object.properties[ref_field.replace("@railConnectionRef", "@atMeasure")]
                except:
                    at_measure = None

                measure_line = MeasureLineString([[x, y, z] for x, y, z in rail_con.geometry.coords])
                measure_result = measure_line.project(imx_object.geometry)
                out_list.append([imx_object.puic, ref_field, ref.imx_object.puic, float(at_measure) if at_measure else None, measure_result.distance_along_line, rail_con.geometry.project(imx_object.geometry)])

df = pd.DataFrame(out_list, columns=["object_puic", "ref_field", "ref_field_value", "imx_measure", "calculated_3d_measure", "calculated_2d_measure"])  #
df.to_excel("measure_check.xlsx", index=False)
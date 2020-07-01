import arcpy


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "PDFExport"
        self.alias = "PDF Export"

        # List of tool classes associated with this toolbox
        self.tools = [BatchPDF]


class BatchPDF(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "AA Batch PDF Export"
        self.description = "Batch Export all PDFs from a selected series"
        self.canRunInBackground = False
        # self.stylesheet = "AADlgContent.xsl"


    def getParameterInfo(self):
        """Define parameter definitions"""
        # Parameter 0
        product_library_con = arcpy.Parameter(
            displayName="Product Library Connection",
            name="product_library_con",
            datatype="DEWorkspace",
            parameterType="Required",
            direction="input")
        #
        # try:
        #     db_con = r"Database Connections\CartoLive_Prodlib.sde"
        #     if not arcpy.Exists(db_con):
        #         raise db_conError
        # except db_conError:
        #     arcpy.AddError("There was an error connecting to the product Library.\n"+arcpy.GetMessages(2))
        # except Exception as e:
        #     arcpy.AddError(arcpy.GetMessages(2))

        # Parameter 1
        solution_name = arcpy.Parameter(
          displayName="Solution",
          name="solution_name",
          datatype="GPString",
          parameterType="Required",
          direction="input")
        solution_name.filter.type = "ValueList"

        # Populate list with solutions from product library
        # try:
        #     result = arcpy.PLListItems_production(product_library_con, "Products")
        #     item_list = result.getOutput(0).split(";")
        #     solution_name.filter.list = item_list
        # except:
        #     arcpy.AddError(arcpy.GetMessages(2))

        # Parameter 2
        product_class_name = arcpy.Parameter(
          displayName="Product Class",
          name="product_class_name",
          datatype="GPString",
          parameterType="Required",
          direction="input")
        product_class_name.filter.type = "ValueList"

        # Parameter 3
        series_name = arcpy.Parameter(
          displayName="Series",
          name="series_name",
          datatype="GPString",
          parameterType="Required",
          direction="input")
        series_name.filter.type = "ValueList"

        # Parameter 4
        annotation_layer = arcpy.Parameter(
            displayName="Choose the annotation layer file for masking",
            name='annotation_layer',
            datatype="DELayer",
            parameterType="Optional",
            direction="input")

        # Parameter 5
        destination_folder = arcpy.Parameter(
          displayName="Choose the destination directory for the PDFs",
          name="destination_folder",
          datatype="DEFolder",
          parameterType="Required",
          direction="input")

        # Parameter 6
        settings_file = arcpy.Parameter(
          displayName="Choose the production settings XML file",
          name="settings_file",
          datatype="DEFile",
          parameterType="Required",
          direction="input")

        # Parameter 7
        all_pages = arcpy.Parameter(
          displayName="All Pages",
          name="all_pages",
          datatype="GPBoolean",
          parameterType="Required",
          direction="input",
          )
        all_pages.value = "True"

        # Parameter 8
        page_range = arcpy.Parameter(
          displayName="Page Range",
          name="page_range",
          datatype="GPString",
          parameterType="Optional",
          direction="input",
          enabled=False)

        # Parameter 9
        pagination_file = arcpy.Parameter(
            displayName="Pagination File",
            name="pagination_file",
            datatype="DEFeatureClass",
            parameterType="Required",
            direction="input"
        )

        return[product_library_con,
               solution_name,
               product_class_name,
               series_name,
               annotation_layer,
               destination_folder,
               settings_file,
               all_pages,
               page_range,
               pagination_file]

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        try:
            if arcpy.CheckExtension("foundation") == "Available":
                # Check out Production Mapping license
                arcpy.CheckOutExtension("foundation")
            else:
                raise Exception
        except:
            return False
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        db_con = r"Database Connections\CartoLive_Prodlib.sde"
        # try:
        #     db_con = r"Database Connections\CartoLive_Prodlib.sde"
        #     if not arcpy.Exists(db_con):
        #         raise db_conError
        # except db_conError:
        #     arcpy.AddError("There was an error connecting to the product Library.\n"+arcpy.GetMessages(2))
        # except Exception as e:
        #     arcpy.AddError(arcpy.GetMessages(2))

        if parameters[0].altered and not parameters[0].hasBeenValidated:
            db_con = parameters[0].value
            result = arcpy.PLListItems_production(db_con, "Products")
            item_list = result.getOutput(0).split(";")
            parameters[1].filter.list = item_list
        else:
            arcpy.AddError(arcpy.GetMessages(2))

        if parameters[1].altered and not parameters[1].hasBeenValidated:
            result = arcpy.PLListItems_production(parameters[0].value, "Products::"+parameters[1].value)
            item_list = result.getOutput(0).split(";")
            parameters[2].filter.list = item_list
        else:
            parameters[2].filter.list = []

        if (parameters[1].altered and parameters[2].altered) and not (parameters[1].hasBeenValidated and parameters[2].hasBeenValidated):
            result = arcpy.PLListItems_production(parameters[0].value, "Products::"+parameters[1].value+"::"+parameters[2].value)
            item_list = result.getOutput(0).split(";")
            parameters[3].filter.list = item_list
        else:
            parameters[3].filter.list = []

        if parameters[7].value:
            parameters[8].enabled = False
        else:
            parameters[8].enabled = True

        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        # -------------------------------------------------------------------------------
        # Name:         Batch PDF Export
        # Purpose:      Traverses the Product Library, checks out the required product
        #               and exports the file to a Production PDF using AA default
        #               settings. The required product is then checked back in to the
        #               Product Library, with an os based clean to catch any copys not
        #               removed during check in
        #
        # Author:      chapmang
        #
        # Created:     15/03/2019
        # Copyright:   (c) chapmang 2019
        # Licence:
        # -------------------------------------------------------------------------------

        import arcpy
        import arcpyproduction
        import os
        import errno
        import re
        import sys
        import time

        # Fetch values from user input
        solution_name = parameters[1].valueAsText
        product_class_name = parameters[2].valueAsText
        series_name = parameters[3].valueAsText
        all_pages = parameters[7].valueAsText
        page_range = parameters[8].valueAsText
        pagination_file = parameters[9].valueAsText

        # If the page range parameter has any value convert string into list individual of values
        # Has to be strings for find replace against contents of product library to work
        if page_range and page_range.strip():
            raw_list = page_range.split(",")
            clean_list = []
            first_product_parts = []
            last_product_parts = []
            for value in raw_list:

                if "-" in value:
                    # A range has been submitted find all the values between start and end
                    split_list = value.split("-")

                    # If the product name contains underscore to separate name from number
                    if "_" in split_list[0]:
                        # The first page split by underscore
                        first_product_parts = split_list[0].split("_")
                        # The last page split by underscore
                        last_product_parts = split_list[-1].split("_")
                        # The first page number (should always be second value in file name, allows for quarters)
                        first_page_number = int(first_product_parts[1])
                        # The last page number plus one to make sure it the whole list is covered
                        # (should always be second value in file name, allows for quarters)
                        last_page_number = int(last_product_parts[1]) + 1

                    else:
                        first_page_number = int(split_list[0])
                        last_page_number = int(split_list[-1]) + 1

                    filled_list = [x for x in range(first_page_number, last_page_number)]
                    # arcpy.AddMessage(filled_list)
                    for a in filled_list:
                        if len(first_product_parts):
                            if a < 1000:
                                # arcpy.AddMessage("Prefix: " + first_product_parts[0])
                                # arcpy.AddMessage("Page: " + str(a).strip())
                                clean_list.append(first_product_parts[0] + "_" + str(a).strip() + "_NW")
                                clean_list.append(first_product_parts[0] + "_" + str(a).strip() + "_NE")
                                clean_list.append(first_product_parts[0] + "_" + str(a).strip() + "_SE")
                                clean_list.append(first_product_parts[0] + "_" + str(a).strip() + "_SW")
                            else:
                                clean_list.append(first_product_parts[0] + "_" + str(a).strip())
                        else:
                            if a < 1000:
                                clean_list.append(str(a).strip() + "_NW")
                                clean_list.append(str(a).strip() + "_NE")
                                clean_list.append(str(a).strip() + "_SE")
                                clean_list.append(str(a).strip() + "_SW")
                            else:
                                clean_list.append(str(a).strip())
            else:
                    page_no = int(value.split("_")[1])
                    if page_no < 1000:
                        # A single product/page number
                        clean_list.append(value.strip() + "_NW")
                        clean_list.append(value.strip() + "_NE")
                        clean_list.append(value.strip() + "_SE")
                        clean_list.append(value.strip() + "_SW")
                    else:
                        clean_list.append(value.strip())
        else:
            product_library_itempath = "Products::" + solution_name + "::" + "::" + product_class_name + "::" + series_name
            result = arcpy.PLListItems_production(r"Database Connections\CartoLive_Prodlib.sde", product_library_itempath)
            clean_list = result.getOutput(0).split(";")

        # SET GLOBAL PARAMETERS
        # Make sure the temporary checkout destination exists
        check_out_path = r"C:/temp/arcpdf/"
        directory = os.path.dirname(check_out_path)

        if not os.path.isdir(directory):
            try:
                os.makedirs(directory)
            except OSError as e:
                if e.errno != errno.EEXIST:
                    raise
        else:
            try:
                # Check to clean up local copies not removed during previous export
                file_list = os.listdir(directory)
                for i in file_list:
                    os.remove(check_out_path + i)
                del file_list
            except Exception as e:
                arcpy.AddMessage(arcpy.GetMessages(0) + 'GEOFF')

        # Get the database connection and check that it valid
        try:
            db_con = r"Database Connections\CartoLive_Prodlib.sde"
            if not arcpy.Exists(db_con):
                raise db_conError
        except db_conError:
            arcpy.AddError("There was an error connecting to the product Library.\n"+arcpy.GetMessages(2))
        except Exception as e:
            arcpy.AddError(arcpy.GetMessages(2))

        # Set output path for finished PDFs
        opath = parameters[5].valueAsText

        # Regular expression for page number check/updating
        regex = re.compile(r""" AND PAGE_NO\s*=\s*[0-9]*""")

        # Set source of Production PDF settings file
        settings_file = settings_file = parameters[6].valueAsText

        # Fetch the child nodes on a give series in the Product library
        # Format = Products::<Solution>::<ProductClass>::<Series>::<Product>::<Instance>::<AOI>
        product_library_itempath = "Products::" + solution_name + "::" + "::" + product_class_name + "::" + series_name

        # List products in the selected series path
        result = arcpy.PLListItems_production(db_con, product_library_itempath)
        item_list = result.getOutput(0).split(";")

        # If range was submitted filter the list to include only products in requested series
        if len(clean_list) > 0 and all_pages == "false":
            filtered_item_list = [x for x in clean_list if x in item_list]
        else:
            filtered_item_list = item_list

        # Loop through the list of products, check each one out, export to PDF and then check it back in
        for i in filtered_item_list:

            # Try each product in turn but don't fail for exception on each one.
            try:
                # print("Product: {0} from Series: GBRA opened".format(i))
                # Concatenate the path to product document for this iteration of the loop
                product_path = product_library_itempath + "::" + str(i) + "::" + str(i) + ".mxd"

                # Check out the product document
                checked_out_file = arcpy.PLCheckoutFile_production(db_con,
                                                                   product_path,
                                                                   check_out_path,
                                                                   product_library_ownername='PRODLIB'
                                                                   )
                arcpy.AddMessage("Product: {0} from Series: {1} checked out".format(i, series_name))

                # Make the Checked Out document active
                current_mxd = arcpy.mapping.MapDocument(os.path.join(check_out_path, i + ".mxd"))

                # Replace definition query page number with mxd number
                filename = os.path.basename(current_mxd.filePath).split(".")[0]
                pageNumber = filename.split("_")[1]

                layers = arcpy.mapping.ListLayers(current_mxd)
                for lyr in layers:
                    if lyr.supports("DEFINITIONQUERY"):
                        # Annotation Layer is classified as Group Layer
                        if lyr.isGroupLayer:
                            lyr.definitionQuery = re.sub(regex, " AND PAGE_NO = " + pageNumber, lyr.definitionQuery)
                        # All other layers are Feature Layers
                        elif lyr.isFeatureLayer:
                            lyr.definitionQuery = re.sub(regex, " AND PAGE_NO = " + pageNumber, lyr.definitionQuery)

                # Re-centre using external pagination file
                df = arcpy.mapping.ListDataFrames(current_mxd)[0]
                lyr = arcpy.MakeFeatureLayer_management(
                    pagination_file,
                    "temp_pagination").getOutput(0)
                arcpy.mapping.AddLayer(df, lyr, "BOTTOM")
                extent_layer = arcpy.mapping.ListLayers(current_mxd, "temp_pagination", df)[0]
                arcpy.SelectLayerByAttribute_management(extent_layer, "NEW_SELECTION", " \"Export_Name\" = \'" + filename + "\' ")
                df.panToExtent(extent_layer.getSelectedExtent(False))

                arcpy.Delete_management(lyr)
                arcpy.mapping.RemoveLayer(df, extent_layer)

                # Set full output path for the exported PDF
                outputPath = os.path.join(opath, i + ".pdf")

                # Export to Production PDF
                arcpyproduction.mapping.ExportToProductionPDF(current_mxd,
                                                              outputPath,
                                                              settings_file,
                                                              data_frame="PAGE_LAYOUT",
                                                              resolution=750,
                                                              image_quality="BEST",
                                                              colorspace="CMYK",
                                                              compress_vectors=True,
                                                              image_compression="LZW",
                                                              picture_symbol="VECTORIZE_BITMAP"
                                                              )
                arcpy.AddMessage("Product: {0} from Series: {1} Exported".format(i, series_name))

                # Save file to allow any definition query or extent corrections to be persisted
                # current_mxd.saveACopy(os.path.join(opath, os.path.basename(current_mxd.filePath)))
                current_mxd.save()

                # Check In product document to release lock in Production Library
                # Doesn't always remove local copy, see clean up below
                if os.path.exists(checked_out_file[0]):
                    overwrite_version = "OVERWRITE_VERSION"
                    keep_checkedout = "DO_NOT_KEEP_CHECKEDOUT"
                    keep_localcopy = "REMOVE_LOCAL_COPY"
                    arcpy.PLCheckinFile_production(db_con, product_path,  overwrite_version, keep_checkedout, keep_localcopy)
                arcpy.AddMessage("Product: {0} from Series: {1} checked in".format(i, series_name))
            except Exception as e:
                arcpy.AddError("product: {0} from Series: {1} failed".format(i, series_name))
                arcpy.AddError("Error on line {} {} {}".format(sys.exc_info()[-1].tb_lineno, type(e).__name__, e))
                arcpy.AddError(e)
                continue
        del item_list

        # Check in the extension
        arcpy.CheckInExtension("foundation")

        return

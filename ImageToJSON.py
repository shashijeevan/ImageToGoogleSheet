# Copyright (C) 2023 Shashi Jeevan Midde Peddanna
#
# This file is released under the "GNU General Public License v3.0". 
# Please see the LICENSE file that should have been included as 
# part of this package.
from RequestDataClasses import RGBColor
from RequestDataClasses import BackgroundColorStyle
from RequestDataClasses import Value
from RequestDataClasses import Row
from RequestDataClasses import Range
from RequestDataClasses import UpdateCellsClass
from RequestDataClasses import Request
from RequestDataClasses import UpdateCells
from RequestDataClasses import UserEnteredFormat
import logging

class RequestBuilder:

    def ImageArrayToJSON(SheetID, image_array, width, height):

        print("Array dimension = {}", image_array.shape)

        logging.info("Array dimension = %s", image_array.shape)

        #print("Width = " + width + ", Height = " + height)

        range1 = Range(sheet_id=SheetID, start_row_index=0, end_row_index=height+1, 
                       start_column_index=0, end_column_index=width+1)

        rows = []

        for rowid in range(0, height):
            
            columns = []

            for columnid in range(0, width):

                #Current Pixel value in RGB
                pixel = image_array[rowid * width + columnid]
                
                pixelcolor = RGBColor(red = pixel[0] / 255, 
                                      green = pixel[1] / 255, 
                                      blue = pixel[2] / 255)

                bgcolor = BackgroundColorStyle(rgb_color=pixelcolor)

                myvalue = Value(user_entered_format=UserEnteredFormat(
                                                        background_color_style=bgcolor))

                #print("Adding Cell to col = {} ".format(columnid))

                logging.info("Adding Cell col = %d of Row = %d", columnid, rowid)

                logging.info(myvalue.to_json())

                columns.append(myvalue)

            #print("Adding row = {}".format(rowid))

            logging.info("Adding row = %d", rowid)

            newrow = Row(values=columns)
            logging.info(newrow.to_json())

            rows.append(newrow)

        cells = UpdateCellsClass(range=range1, rows=rows, fields='UserEnteredFormat')
                            
        #print(rows)

        request = Request(update_cells=cells)

        updaterequest = UpdateCells(requests=[request])

        result = updaterequest.to_json()

        return result
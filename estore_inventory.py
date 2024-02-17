# -*- coding: utf-8 -*-
"""
Created on Tue Dec  5 12:22:59 2023

Manage inventory of an ebay store, classified by category and with linked
images.

estore.db is created by loading estore.xlsx into sqlite using estoredb.py .

STATUS: STILL WORKING ON VALIDATION OF ITEM EDIT ... 
    ... if quit button is hit when item editing is in progress with changes (itemChanged) then quit should ask whether to save current changes.
          
    
    TODO: In delete of image, imgThumbs list needs to have image removed, and "row" number adjust accordingly.
          Should also probably prompt to confirm deletion, and flag if Primary image is deleted, maybe prompt if primary is about to be deleted.
          
          ... rotation does not work consistently maybe the rotation coming out of the df is not updated when the db rotation changes, or maybe
              does not update when rotation goes from 270 to 360 / 0 - see first image of item 4542.
    
    TODO: Image zoom, if click in image then < > / mouse left-right keys to scroll through images.
    
    TODO: Check on parent.dB.txtQw heirarchy in image deletion.
    
    TODO: Image thumbnails need a specific border to highlight currently displayed image, that also needs to allow current+primary = same
        
        add tooltip hints to combobox list, with full description of values.

          
@author: volke
"""

import estore_pptx
from estore_pptx import pptxCatalog

from PyQt6 import QtCore, QtWidgets, QtGui
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, 
    QWidget, 
    QDockWidget,
    QGroupBox,
    QHBoxLayout,
    QVBoxLayout,
    QGridLayout,
    QPushButton,
    QLabel,
    QTextEdit,
    QLineEdit,
    # QPlainTextEdit,
    QComboBox,
    QFileDialog,
    QAbstractItemView,
    QAbstractItemDelegate,
    QStyledItemDelegate,
    QProgressBar,
    QProgressDialog,
    QHeaderView,
    QTableWidget, 
    QTableWidgetItem,
    QTableView,
    QSizePolicy,
    QMenuBar,
    QToolBar,
    QMenu
    )

from PyQt6.QtCore import (
    QSize, Qt, QTimer, QDateTime, 
    QDir,
    QEvent,
    QAbstractTableModel,
    QAbstractItemModel,
    QSortFilterProxyModel,
    QModelIndex
    )


from PyQt6.QtGui  import (
    QPixmap,
    QMouseEvent,
    QAction,
    QIcon,
    QPalette,
    QColor,
    QStandardItem,
    QTransform,
    QFontMetrics,
    QRegularExpressionValidator, QIntValidator, QDoubleValidator
    )

from functools import partial

from PIL import Image
from datetime import date
import subprocess
import sqlite3
import os.path
import re
import sys
import shutil
import pandas as pd
  
#----------------------------------------------------------------------
# GLOBAL SETTINGS

ESIfontSize          = 10
ESItimerResize       = 600  # Timeout delay in ms after resize before repaint
ESItimerFilterText   = 500  # Timeout delay in ms after last keystroke before filter refresh
ESItimerWarnStyle    = 1500 # Timeout delay for warning color when invalid characters typed in a field.
ESIimgthmStyleCat    = "border:0px solid #7dafc1; background-color: white;"
ESIimgthmStylePmy    = "border:0px solid #00ff00; background-color: #d0ffd0;"
ESIimgthmStyleCur    = "border:0px solid #ff3030; background-color: white;"
ESIimgthmStyleKeep   = "border:0px solid #a0a0a0; background-color: #404040;"
ESIupperCaptionStyle = "font: italic 10pt; alignment: bottom right; color: #505070;"
ESIinputFieldOkStyle = "background-color: #e8ffff;"
ESIinputFieldWarnStyle = "background-color: #f7cafa;"
ESIalignTopRight     = QtCore.Qt.AlignmentFlag.AlignRight | QtCore.Qt.AlignmentFlag.AlignTop
ESIalignTopLeft      = QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignTop
ESIalignBottomLeft   = QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignBottom
ESIcaptionFont       = "font: italic %dpt; color:#404040; " % (ESIfontSize)
ESIcaptionFontColor  = "#0a28FF"

ESItextLineHeight    = 20 # Typical label height for a single line label
ESIcharacterWidth    = 7.5 * ESIfontSize / 10.0  # Typical text character width in pixels
ESIlistColPadding    = 10    # Additional width of each item list column beyond data 
ESIimgthmWidth       = 80
ESIimgthmSize        = QSize( ESIimgthmWidth, ESIimgthmWidth )
ESIimgPptxDpi        = 600   # Target dots per inch for pptx images
ESIimageListSize     = QSize( ESIimgthmWidth+10, 800 )
ESIbuttonMaxSize     = QSize(    80,    ESItextLineHeight )
ESIimageMinSize      = QSize(   100,   350 )
ESIimageMaxSize      = QSize(  3000,  2000 )
ESIimageDefaultSize  = QSize(   800,   800 )
ESIitemLabelMinWidth = 80
ESIitemValueMinWidth = 400
ESIitemPanelMinWidth = ESIitemLabelMinWidth + ESIitemValueMinWidth
ESIitemListMinWidth  = min( 700, ESIitemPanelMinWidth )
ESIitemListMinSize   = QSize( ESIitemListMinWidth,   100 )
ESItextWinMinWidth   = min( 700, ESIitemPanelMinWidth )
ESItextWinMinSize    = QSize( ESItextWinMinWidth,   100 )
ESIdataGroupMaxSize  = QSize( 2000, 2000 )

ESIimgSizePolicy = QSizePolicy( 
        QSizePolicy.Policy.Expanding,
        QSizePolicy.Policy.Expanding )
ESIimgSizePolicy.setHorizontalStretch( 5 )
ESIimgSizePolicy.setVerticalStretch( 5 )

ESIimlSizePolicy = QSizePolicy( 
        QSizePolicy.Policy.Fixed,
        QSizePolicy.Policy.Preferred )
ESIimlSizePolicy.setHorizontalStretch( 1 )
ESIimlSizePolicy.setVerticalStretch( 5 )

ESIdtSizePolicy = QSizePolicy( 
                QSizePolicy.Policy.Expanding,
                QSizePolicy.Policy.Preferred )
ESIdtSizePolicy.setHorizontalStretch( 10 )
ESIdtSizePolicy.setVerticalStretch( 1 )

ESIdtlSizePolicy = QSizePolicy(                   # Item panel left hand labels
                QSizePolicy.Policy.Preferred,
                QSizePolicy.Policy.Preferred )
ESIdtlSizePolicy.setHorizontalStretch( 1 )
ESIdtlSizePolicy.setVerticalStretch( 1 )

ESIdtdSizePolicy = QSizePolicy(                   # Item panel right hand value fields
                QSizePolicy.Policy.Expanding,
                QSizePolicy.Policy.Preferred )
ESIdtdSizePolicy.setHorizontalStretch( 5 )        
ESIdtdSizePolicy.setVerticalStretch( 1 )

ESIdttSizePolicy = QSizePolicy(                   # Item panel right hand value text edit fields
                QSizePolicy.Policy.Expanding,
                QSizePolicy.Policy.Expanding )
ESIdttSizePolicy.setHorizontalStretch( 5 )        
ESIdttSizePolicy.setVerticalStretch( 5 )

ESIitmListSizePolicy = QSizePolicy( 
                QSizePolicy.Policy.Expanding,
                QSizePolicy.Policy.Expanding )
ESIitmListSizePolicy.setVerticalStretch( 3 )
ESIitmListSizePolicy.setHorizontalStretch( 10 )


#----------------------------------------------------------------------
"""
Frame Layout

 +-apL------------------------------+--------------------------------------+
 | (dtL)                            | (imL)                                |
 | +-itmEW------------------------+ | +---------+ +----------------------+ |
 | | item editing panel (itmEDL)  | | | imgList | | imgView              | |
 | | +--------------------------+ | | | (imLF)  | | (imVF)               | |
 | | | edit buttons (itmEBL)    | | | |         | |                      | |
 | | +--------------------------+ | | |         | |                      | |
 | +------------------------------+ | |         | |                      | |
 | +-itmLL------------------------+ | |         | |                      | |
 | | itemList (itmLW)             | | |         | |                      | |
 | +------------------------------+ | |         | |                      | |
 | +-monL-------------------------+ | |         | |                      | |
 | | text message console (txtQW) | | |         | |                      | |
 | +------------------------------+ | |         | |                      | |
 |                                  | |         | +----------------------+ |
 |                                  | |         | + imgInfo (imIF)       + |
 | +-appBL------------------------+ | |         | +----------------------+ |
 | | App  control buttons         | | |         | | Photo buttons (imBF) | |
 | +------------------------------+ | +---------+ +----------------------+ |
 +----------------------------------+--------------------------------------+

 """
   
#---------------------------------------------------------------------------------
# Data management class - dataframes are used for current selection of data only,
# edited and inserted rows are updated in the db immediately, ie. the dataframe
# should not be used to replace table contents.
# Column structure should be defined in the original estore.xlsx spreadsheet and
# translated to sqlite using estoredb.py 

class itemDb:
    
    def __init__( self, dbPath, category, txtQW ):

        self.itpKey    = { "Field": "ItemID", "Label": "ID", "Width": 8, "Type": str, "Widget": None }
        
        self.con = sqlite3.connect( dbPath )
        self.dbPath    = dbPath
        self.cur = self.con.cursor()
        self.category  = category
        self.Album     = None
        self.itmDf     = None
        self.imgDf     = None
        self.keyCol    = self.itpKey["Field"]
        self.txtQW     = txtQW
        self.lookups   = {}
        self.loadFields()
        self.refresh()

    def defaultDf( self ): # create dataframe with one row of default values
        defaultFrm = {}
        for cL in self.itpFields:
            defaultFrm[cL["Field"]] = [cL["Default"]]
        return pd.DataFrame( defaultFrm )

    def loadFields( self ): # Get configuration of item fields specific to this category
        # Populate from ItemFields with selection and modifiers based on CategoryFields 
        catFd = pd.read_sql_query( 'select * from CategoryFields where Category = \"%s\" ORDER BY Seq ASC' %(self.category), self.con )
        for dbCol in [ 'index', 'Category', 'Seq' ]:
            catFd = catFd.drop( dbCol, axis=1 )
        catFd.replace( [None], "", inplace = True )

        itpNew = []
        for cF in catFd.index:
            catInfo = {}
            field = catFd["Field"][cF]
            for cI in catFd.columns:
                catInfo[cI] = catFd[cI][cF]

            selQuery = 'select * from ItemFields where Field = \'%s\'' %(field)
            fieldFd = pd.read_sql_query( selQuery, self.con )
            fieldFd = fieldFd.drop( 'index', axis = 1)
            fieldFd.replace( [None], "", inplace = True )

            fieldInfo = {}
            for col in fieldFd.columns:
                fieldInfo[col] = fieldFd[col][0]
            for uiCol in ['UI', 'Label', 'Width', 'Type', 'Format', 'Default', 'ListVisible', 'PanelVisible', 'Validator', 'Validator1', 'Validator2' ]:
                if not catInfo[uiCol] == "":
                    fieldInfo[uiCol] = catInfo[uiCol]
            cL = { "Widget": None,
                   "Edited": False,
                   "Width": 20,
                   "Default": "",
                   "Valid": True,
                   "Validator": None,
                   "ListVisible": 2,
                   "PanelVisible": 2,
                   "Filters": None }
            for cLf in fieldInfo:
                cL[cLf] = fieldInfo[cLf]
            cL["Value"] = cL["Default"]
            validator = cL["Validator"]
            if validator == 'Lookup':
                cL['Validator'] = ( validator, fieldInfo['Validator1'])
            elif validator == 'Range':
                cL['Validator'] = ( validator, [fieldInfo['Validator1'],fieldInfo['Validator2']])
            else:
                cL['Validator'] = None
            
            # ListVisible controls presence in item list - 0 not 1 visible but not filtered 2 visible and filtered
            # PanelVisible controls presence in item panel - 0 not 1 visible but not editeable 2 visible and editable
            # Put all items in the cL anyway as new items that are hidden will still get non-blank default values assigned.

            itpNew.append( cL )

        self.itpFields = itpNew
        catSql = "Select AlbumID,ItmPrefix,ItmSequence,ImgPrefix,ImgSequence,Description from CATEGORIES where Category = \"%s\"" %(self.category)
            
        res = self.cur.execute( catSql )
        catInfo = res.fetchone()
        if catInfo:
            cI = 0
            self.Album            = catInfo[cI]
            cI = cI + 1
            self.ItmPrefix        = catInfo[cI]
            cI = cI + 1
            self.ItmSequence      = catInfo[cI]
            cI = cI + 1
            self.ImgPrefix        = catInfo[cI]
            cI = cI + 1
            self.ImgSequence      = catInfo[cI]
            cI = cI + 1
            self.Description      = catInfo[cI]
        else:
            print("ERROR: CATEGORIES table not found as expected in database %s" %( self.dbPath ))
        albumSql = "Select Path from ALBUMS where AlbumID = \"%s\"" %( self.Album ) 
        res = self.cur.execute( albumSql )
        albumInfo = res.fetchone()
        if albumInfo:
            self.albumPath        = albumInfo[0]
            if os.name == 'nt':
                self.albumPath = re.sub( r"/home/volker/", r"c:\\Users\\volke\\", self.albumPath )
                self.albumPath = re.sub( r"/", r"\\", self.albumPath )

            if os.path.isdir( self.albumPath):
                self.txtQW.insert( "Using %s as path for images." %(self.albumPath))
            else:
                self.txtQW.insert("ERROR: Album path %s is not a valid folder" %(self.albumPath ))
        else:
            self.txtQW.insert("ERROR: Album path not found in database %s" %( self.dbPath ))


    def columnFilters( self, dbField, filterType, filters):
        cL = self.itpFields[dbField]
        cL["Filters"] = ( filterType, filters )

    def refresh( self ):

        itmSql = "SELECT * from ITEMS where Category = \'%s\'" %(self.category)
        sqlAnd = " AND "
        for cL in self.itpFields:
            if not cL["Filters"] is None:
                if cL["Filters"][0] == "List":
                    if len( cL["Filters"][1] ) > 0:
                        itmSql = "%s%s\"%s\" IN (" % (itmSql, sqlAnd, cL["Field"])
                        sqlComma = ""
                        for cFilter in cL["Filters"][1]:
                            itmSql = "%s%s'%s'" % ( itmSql,sqlComma, cFilter )
                            sqlAnd = " AND "
                            sqlComma = ","
                        itmSql = itmSql + ")"
                elif cL["Filters"][0] == "Text":
                    if not cL["Filters"][1][0] == "":
                        wildCard = '\'%' + cL["Filters"][1][0] + '%\''
                        itmSql = "%s%s\"%s\" LIKE (%s)" % (itmSql, sqlAnd, cL["Field"], wildCard )
                elif cL["Filters"][0] == "Float":  # Should actually retain the 'float' filter type and rename this as 'Range'
                    if not cL["Filters"][1][0] is None:
                        fromValue = cL["Filters"][1][0]
                        toValue   = cL["Filters"][1][1]
                        itmSql = "%s%s\"%s\" BETWEEN %f and %f" % (itmSql, sqlAnd, cL["Field"], fromValue, toValue )
        self.txtQW.insert( "Reloading table with query - %s" % (itmSql))
        self.itmDf = pd.read_sql_query( itmSql, self.con )
        
        # Replace any missing values, check that all index values are appropriate
        
        self.itmDf[self.keyCol] = self.itmDf[self.keyCol].astype('string')
        for cL in self.itpFields:
            column = cL["Field"]
            dType  = cL["Type"]
            if dType =="float":
                self.itmDf[column].fillna( 0.0, inplace = True )
            elif dType == "int":
                self.itmDf[column].fillna( 0, inplace = True )
                self.itmDf[column] = self.itmDf[column].apply(int)
            else:
                self.itmDf[column].replace( [ None ], "", inplace = True )
        blankKeys = self.itmDf.loc[self.itmDf[self.keyCol].isna() ]
        blankCount = len(blankKeys)
        if blankCount > 0:
            self.txtQW.insert("ERROR: Blank ItemID rows found in database on rows:")
            
            for index, row in blankKeys.iterrows():
                self.txtQW.insert( "- Index %s, Name %s" % (index, row['Name']) ) 
            
            self.itmDf = self.itmDf.drop(self.itmDf[self.itmDf[self.keyCol].isna()].index)
            self.itmDf.reset_index( inplace = True )

        # Get lookups as dicts with keys
        #  ID      value stored in db
        #  Label   value shown in user interface
        #  Text    tooltip explanation for user interface
        
        getLookups = "select distinct LOOKUP from LOOKUPS;"
        res = self.cur.execute( getLookups )
        for resL in res:
            lookup = resL[0]
            # self.lookups[lookup] = pd.read_sql_query( 'select \'Id\', \'Label\', \'Text\' from LOOKUPS where LOOKUP = \'%s\' order by SEQ ASC' %(lookup), self.con )
            self.lookups[lookup] = pd.read_sql_query( 'select ID, Label, Text from LOOKUPS where LOOKUP = \"%s\" order by SEQ ASC' %(lookup), self.con )
            
        # self.txtQW.insert( "DB connection refreshed.")
        
    def nextItemID( self ):
        self.ItmSequence += 1
        updateSql = "update CATEGORIES set ItmSequence=%d where Category = \'%s\'" % ( self.ItmSequence, self.category )
        try: 
            self.cur.execute( updateSql )
            self.con.commit( )
        except:
            self.txtQW.insert( "DB error updating item sequence number" )    
        return "%s%s" % (self.ItmPrefix, self.ItmSequence)

    def nextImageFile( self, itemId, incomingFilePath ):
            # incomingFilePath is path of file being added which will be copied into the album;
            # if none a new file name is generated but no file actions are undertaken.

        self.ImgSequence = self.ImgSequence + 1
        picId   = "%s_%d" % (self.ImgPrefix, self.ImgSequence )
        picPath = None
        if not incomingFilePath is None:
            if not os.path.exists( incomingFilePath ):
                self.txtQW.insert( "Proposed new image file %s not found.")
                picPath = None
            else:
                newFileParts = os.path.splitext( incomingFilePath )  # get extension
                picExt  = newFileParts[1].lower()
                picFile = "%s%s" % (picId, picExt )
                picPath = "%s%s" % ( self.albumPath, picFile )
                try:
                    shutil.copy( incomingFilePath, picPath )
                    self.txtQW.insert( "Image file copied as %s" % ( picPath ))
                except:
                    self.txtQW.insert( "Error on image file copy from %s to %s" % ( incomingFilePath, picPath ))    
                    picPath = None
        else:
            picFile = "%s.jpg" % (picId)
            picPath = "%s%s" % (self.albumPath, picFile )
            
        if not picPath is None:

            dateAdded = date.today().strftime( "%Y%m%d")
            insertSql = """INSERT INTO IMAGES (
                AlbumID, ImageID, ItemID, \'Rank\', Zoomed, Edited, Rotation, DateAdded, Copyright )
                VALUES
                ( \'%s\', \'%s\', \'%s\', \'%s\', \'%s\', \'%s\', %d, \'%s\', \'%s\' );
                 """ % (
                 self.Album, picId, itemId, "C", "", "", 90, dateAdded,  'C_VH' )
        
            self.txtQW.insert( "Insert: %s" %( insertSql))
            try:
                self.cur.execute( insertSql ) 
                self.con.commit()
            except:
                self.txtQW.insert( "ERROR ON DB INSERT OF IMAGE %s as %s for item %s" %(incomingFilePath,picId, itemId) )

            updateSql = "update CATEGORIES set ImgSequence=%d where Category = \'%s\'" % ( self.ImgSequence, self.category )
            try:
                self.cur.execute( updateSql ) 
                self.con.commit()
            except:
                self.txtQW.insert( "ERROR ON DB UPDATE OF IMAGE COUNT IN CATEGORIES" )
                picFile = ""
        else:
            picFile = ""
        return ( picId, picFile, picPath )
    
    def close( self ):
        self.con.close()
        
class imageThumb( QPushButton): # One thumbnail

    def __init__( self, img, imList ):
        super(imageThumb,self).__init__()
        self.img  = img
        self.imView = imList.imView
        self.imList = imList
        self.setMinimumSize( ESIimgthmSize )

        self.resetStyle( )

        if not img["Exists"]:
            self.setToolTip( "%s - MISSING IMAGE" % (img["ImageID"]) )
            self.setEnabled( False )
            img["Thumb"] = None
        else:
            self.setToolTip( img["ImageID"] )
            self.refresh( img )
            self.setEnabled( True )
            img["Thumb"] = self
            self.clicked.connect( partial( self.imList.refresh, self.img["Row"] ))
            
    def refresh( self, img):
        iconPix = img["Pixmap"].scaled(
            ESIimgthmSize,
            aspectRatioMode = QtCore.Qt.AspectRatioMode.KeepAspectRatio)
            # mode = QtWidgets.QGraphicsPixmapItem.SmoothTransformation )
        self.setIcon( QIcon(iconPix) ) # plImage )
        self.setIconSize( ESIimgthmSize )       

    def resetStyle( self ):
        if self.img["Rank"] == 'P':
            self.setStyleSheet( ESIimgthmStylePmy )
        elif self.img["Rank"] == 'K':
            self.setStyleSheet( ESIimgthmStyleKeep )
        else:
            self.setStyleSheet( ESIimgthmStyleCat )       
        
class imageList( QGroupBox ):   # Scrollable list of images

    def __init__( self, imView, parent ):
        
        super().__init__( parent )

        self.imgThumbs = []
        self.imView    = imView
        self.parent    = parent
        self.itemId    = None
        self.imgRows   = QVBoxLayout()
        self.pmyRow    = -1
        self.imgRows.setAlignment( QtCore.Qt.AlignmentFlag.AlignTop )
        self.imgAddButton()
        self.setLayout( self.imgRows )

    def imgAddButton( self ):
        imgAddButton = QPushButton(parent=self.parent, text="Add Image" )
        imgAddButton.setMaximumSize( ESIimgthmSize )
        imgAddButton.clicked.connect( self.imageAdd )
        imgAddButton.setToolTip( "Add another image to current item" )
        imgAddButton.setEnabled( True )
        self.imgRows.addWidget( imgAddButton )

    def imageAdd( self ):
        print("Adding images")
        newItem = False
        iDlg = QFileDialog( self )
        iDlg.setFileMode( QFileDialog.FileMode.ExistingFiles )
        #imgFileInfo = iDlg.getOpenFileName( self, 
        #                        "Open image", 
        #                        self.parent.dB.albumPath,
        #                        "Image files (*.png *.jpg)")
        #imgFile = imgFileInfo[0]
        iDlg.setDirectory( self.parent.dB.albumPath )
        iDlg.setNameFilter( "Images (*.jpg *.png)" )
        iDlg.setViewMode( QFileDialog.ViewMode.Detail )
        if iDlg.exec():
            files = iDlg.selectedFiles()
        else:
            files = []
        for addFile in files:  # First add the image into the images index
                               # Later (if newItem) add the new item
            picId, picFile, picPath = self.parent.dB.nextImageFile( self.itemId ,addFile )
           
        self.parent.itmLW.loadImages( self.itemId, self.parent.itmLW.getImages( self.itemId, False ) )
        
    def loadImages( self, itemId, addImages ):
        self.itemId = itemId
        for iThm in self.imgThumbs:
            iThm.destroy()
        for i in reversed(range(self.imgRows.count())): 
            self.imgRows.itemAt(i).widget().setParent(None)
        row = 0
        self.imgThumbs = []
        imShown = False
        self.imgAddButton()
        if len( addImages ) <= 0:
            primaryImg = self.imView.defaultImg
            self.imView.refresh( self.imView.defaultImg )
        else:
            defaultPmy = True
            primaryImg = addImages[0]
            for img in addImages:
                thm = imageThumb( img, self )

                if img["Rank"] == "P":
                    defaultPmy = False
                    primaryImg = img
                    self.pmyRow = row
                row += 1
                self.imgRows.addWidget( thm )
                self.imgThumbs.append( thm )
            if defaultPmy:
                primaryImg["Rank"] = "P"
                self.pmyRow = 0
            self.imView.refresh( primaryImg )
            # self.refresh( row )

    def refresh( self, row ): 
        
        if row < 0 :
            row = 0
        elif row >= len( self.imgThumbs ):
            row = len( self.imgThumbs ) - 1

        for imgT in self.imgThumbs:
            imgT.resetStyle()               
        self.imgThumbs[row].setStyleSheet( ESIimgthmStyleCur )
        self.imView.refresh( self.imgThumbs[row].img )
        
    def scrollImages( self, rows ):  # Scroll across a number of rows
        print("A")
        
    def switchPrimary( self, row ):  # Change primary image to the one in row
        if row == self.pmyRow:
            self.parent.txtQW.insert( "Image in row %d is already primary." % (row))
        else:
            self.parent.txtQW.insert( "Switch primary from %s (row %d) to %s (row %d)" % (self.imgThumbs[self.pmyRow].img["ImageID"],self.pmyRow, self.imgThumbs[row].img["ImageID"],row))
            updateSql = "update IMAGES set RANK = \'P\' where IMAGEID = \'%s\'" % ( self.imgThumbs[row].img["ImageID"] )
            self.parent.dB.cur.execute( updateSql )
            self.parent.dB.con.commit()
            self.parent.dB.txtQW.insert( updateSql )
            self.imgThumbs[row].img["Rank"] = "P"
            updateSql = "update IMAGES set RANK = \'C\' where IMAGEID = \'%s\'" % ( self.imgThumbs[self.pmyRow].img["ImageID"] )
            self.parent.dB.cur.execute( updateSql )
            self.parent.dB.con.commit()
            self.parent.dB.txtQW.insert( updateSql )
            self.imgThumbs[self.pmyRow].img["Rank"] = "C"
            self.pmyRow = row
            self.refresh( row )
            
    def deleteImage( self, row ):
        img = self.imgThumbs[row].img
        self.parent.txtQW.insert( "Deleting image %s" % (img["Path"]))
        try:
            os.remove( img["Path"] )
            try:
                updateSql = "update IMAGES set RANK=\'D\' where IMAGEID = \'%s\'" % ( img["ImageID"] )
                self.parent.dB.cur.execute( updateSql )
                self.parent.dB.con.commit()
                self.parent.dB.txtQW.insert( updateSql )
                # REMOVE FROM imgThumbs ARRAY - HAVE TO ADJUST ALL ROW ENTRIES
                self.refresh( 0 )
            except:
                self.parent.dB.txtQW.insert( "Deleted image file %s but database update failed." % (img["Path"]))
        except:
            self.parent.dB.txtQW.insert( "Could not delete image %s" % (img["Path"]))

class imageView( QGroupBox ):  # View an image

    def __init__( self, dB, parent ):
        super().__init__( parent )
        self.a = "b"
        self.picImage = None
        self.picInfo  = None
        self.picOK    = False
        self.parent   = parent
        self.scrollCounter = 0
        self.scrollDelta   = 0
        # ( image file name, rotation, availability flag, pixmap, path, rank flag )
        self.img  = {   "ImageID":     "Default image",
                        "Rotation":    0.0,
                        "Exists":      False,
                        "Pixmap":      None,
                        "Thumb":       None,
                        "Path":        "",
                        "Rank":     "" }
        self.defaultPix = QPixmap("/home/volker/Dropbox/private_python/test.jpg")
        self.defaultImg  = {   
                        "ImageID":     "Default",
                        "Rotation":    0.0,
                        "Exists":      False,
                        "Pixmap":      self.defaultPix,
                        "Thumb":       None,
                        "Path":        "Default",
                        "Rank":     "" }

        # self.dbRotn = 0.0 attempt to have rotation not updated in db immediately
        self.dB   = dB
        self.imRows = QVBoxLayout()
        self.imRows.setAlignment( QtCore.Qt.AlignmentFlag.AlignTop )

        # Image Viewing Label
        self.picImage = QLabel( parent=self, text="" )
        self.picImage.setPixmap( self.defaultPix.scaled(
                ESIimageDefaultSize,
                aspectRatioMode = QtCore.Qt.AspectRatioMode.KeepAspectRatio)
                )

        self.picImage.setMinimumSize( QSize( 1, 1 ) ) 
        self.picImage.setMaximumSize( ESIimageMaxSize ) 
        self.picImage.setStyleSheet( "background-color: orange") 
        
        # Image Information Label
        self.picInfo  = QLabel( parent=self, text= "No image selected." )
        self.picInfo.setStyleSheet( "background-color: white") 
        self.picInfo.setFixedHeight( ESItextLineHeight )
        
        # Image Manipulation Buttons

        self.setAlignment( QtCore.Qt.AlignmentFlag.AlignTop )
        self.btnRow = QHBoxLayout()
        self.rotnButton  = QPushButton(parent=self, text="Rotate")
        self.rotnButton.setMaximumSize( ESIbuttonMaxSize )
        self.rotnButton.clicked.connect( self.rotate )
        self.rotnButton.setEnabled( False )

        self.imDelButton  = QPushButton(parent=self, text="Delete")
        self.imDelButton.setMaximumSize( ESIbuttonMaxSize )
        self.imDelButton.clicked.connect( self.imgDelete )
        self.imDelButton.setEnabled( False )

        self.imPmyButton  = QPushButton(parent=self, text="Primary")
        self.imPmyButton.setMaximumSize( ESIbuttonMaxSize )
        self.imPmyButton.clicked.connect( self.imgPrimary )
        self.imPmyButton.setEnabled( False )

        self.btnRow.setAlignment( QtCore.Qt.AlignmentFlag.AlignLeft )

        self.imRows.addWidget( self.picImage )
        self.imRows.addWidget( self.picInfo )
        self.btnRow.addWidget( self.rotnButton )
        self.btnRow.addWidget( self.imDelButton )
        self.btnRow.addWidget( self.imPmyButton )
        self.imRows.addLayout( self.btnRow )

        self.thumbNail = None
        self.setLayout( self.imRows )
        # self.refresh( self.defaultImg )
        
    def updateImageRotation( self ):
        updateSql = "update IMAGES set ROTATION=%.0f where IMAGEID = \'%s\'" % ( self.img["Rotation"], self.img["ImageID"] )
        self.dB.cur.execute( updateSql )
        self.dB.con.commit()
        self.dB.txtQW.insert( updateSql )
        self.dbRotn = self.img["Rotation"]

        
    def rotate( self ):
        if self.picOK:
            rotation = self.img["Rotation"]
            rotation += 90.0
            if rotation >= 360.0:
                rotation = 0.0
            self.img["Rotation"] = rotation
            self.img["Pixmap"] = self.img["Pixmap"].transformed(QTransform().rotate(rotation))
            self.updateImageRotation( )
            self.refresh( self.img )
        
    def imgDelete( self ):
        self.parent.imgList.deleteImage( self.img["Row"] ) 
        """
        if self.picOK:
            self.dB.txtQW.insert( "Deleting image %s" % (self.img["Path"]))
            try:
                os.remove( self.img["Path"] )
                try:
                    updateSql = "update IMAGES set RANK=\"D\" where IMAGEID = \"%s\"" % ( self.img["ImageID"] )
                    self.dB.cur.execute( updateSql )
                    self.dB.con.commit()
                    self.dB.txtQW.insert( updateSql )
                except:
                    self.dB.txtQW.insert( "Deleted image file %s but database update failed." % (self.img["Path"]))
            except:
                self.dB.txtQW.insert( "Could not delete image %s" % (self.img["Path"]))
        """
        
    def imgPrimary( self ):
        if self.picOK:
            self.dB.txtQW.insert( "Making image %s [%s] the primary." % (self.img["Path"], self.img["Rank"]))
            self.parent.imgList.switchPrimary( self.img["Row"] )
            
    def wheelEvent( self, event ):
        if event.angleDelta().y() < 0:
            self.scrollDelta += 1
        else:
            self.scrollDelta -= 1
        if self.scrollCounter == 0:
            # Swap to next image
            # self.txtQW.insert( "Scroll timer restarted.")  # This can probably be reduced to a single call to singleShot from scrollTimer itself.
            self.sTimer = QtCore.QTimer.singleShot( 500, self.scrollTimer )
            self.scrollCounter = 1
        elif self.scrollCounter == 1:
            self.scrollCounter = 2   # This tells the timer event that there have been more resize callbacks since being triggered.
        QtWidgets.QMainWindow.wheelEvent( self, event )

    def scrollTimer( self ):  

        if self.scrollCounter == 1:
            # self.dB.txtQW.insert( "scroll complete - now redraw the image.")
            row = self.img["Row"] + self.scrollDelta
            self.parent.imgList.refresh( row )
            self.scrollCounter = 0
            self.scrollDelta   = 0
        else:
            # self.dB.txtQW.insert( "scroll timer restarted in callback.")
            self.sTimer = QtCore.QTimer.singleShot( 600, self.scrollTimer )
            self.scrollCounter = 1

    def refresh( self, img ):  # ImageView refresh
    
        imgExists = img["Exists"]
        if 1: # img == self.img:
            imgSize   = self.picImage.size()
        else:
            imgSize   = ESIimageDefaultSize # img["Pixmap"].size() 
            # Need a refresh after resize to allow resampling at larger / smaller sizes
            # the current setting locks image size
        if imgExists:
            imgPath = img["Path"]
            viewPix = img["Pixmap"]
            viewPix = viewPix.scaled(
                imgSize,
                aspectRatioMode = QtCore.Qt.AspectRatioMode.KeepAspectRatio)
                # mode = QtWidgets.QGraphicsPixmapItem.SmoothTransformation )
            self.picImage.setPixmap( viewPix )
            # self.picImage.setScaledContents( True )
            rankFlag = " [%s]" % (img["Rank"])
            picText = "File: %s%s" % (imgPath, rankFlag)
            if img["Rotation"] != 0.0:
                picText = picText + " (Rotated %.1f)" % (img["Rotation"])
            self.picInfo.setText( picText )

            self.picOK = True
            self.rotnButton.setEnabled( True )
            self.imDelButton.setEnabled( True )
            self.imPmyButton.setEnabled( True )

        else:
            self.picImage.setPixmap( self.defaultImg["Pixmap"].scaled(
                imgSize,
                aspectRatioMode = QtCore.Qt.AspectRatioMode.KeepAspectRatio))
            if img["Path"] == "Default":
                self.picInfo.setText( "No images selected." )
            else:
                self.picInfo.setText( "File: %s NOT FOUND" % (img["Path"]))
            self.picOK = False
            self.rotnButton.setEnabled( False )
            self.imDelButton.setEnabled( False )
            self.imPmyButton.setEnabled( False )
        self.img      = img
        self.picInfo.setStyleSheet( ESIcaptionFont )
         
        if img["Thumb"] is not None:
            img["Thumb"].refresh( img )
            
class itemPanel( QGroupBox ):   # Item details list and edit panel

    def __init__( self, dB, parent ):
        super().__init__( parent )
        self.dB = dB
        self.txtQW = self.dB.txtQW
        self.itpKey = self.dB.itpKey
        self.itpFields = dB.itpFields
        self.indexVal  = None
        self.itmLW = None
        self.itemEdited = False

        ipRows = QVBoxLayout()
        ipVL   = QGridLayout()
        ipBL   = QHBoxLayout()
        ipBL.setAlignment( QtCore.Qt.AlignmentFlag.AlignLeft )
        iLb= QLabel( parent=self, text="ID")
        iLv= QLabel( parent=self, text=self.itpKey["Label"])
        iLb.setStyleSheet( ESIcaptionFont )
        iLb.setAlignment( ESIalignTopRight )
        iLv.setAlignment( ESIalignTopLeft )
        self.itpKey["Widget"] = iLv
        
        ipVL.addWidget( iLb, 0,0 )
        ipVL.addWidget( iLv,0,1 )
        ipVL.setColumnMinimumWidth( 0, ESIitemLabelMinWidth )
        ipVL.setColumnMinimumWidth( 1, ESIitemValueMinWidth )

        rowNo = 1
        for cL in self.itpFields:
            if not cL['PanelVisible'] == 0:
                iLb = QLabel( parent=self, text=cL["Label"] )
                iLb.setStyleSheet( ESIcaptionFont )
                iLb.setAlignment( ESIalignTopRight )
                iLb.setSizePolicy( ESIdtlSizePolicy )
                itpRow = self.dB.defaultDf() # self.dB.itmDf[self.dB.itmDf[self.itpKey["Field"]] == self.indexVal ]
                itpValue = itpRow[cL["Field"]].values[0]
                cL["LWidget"] = QLabel( parent=self, text=str(itpValue) , width=cL["Width"] * ESIcharacterWidth )
                cL["LWidget"].setWordWrap( True )
                cL["LWidget"].setSizePolicy( ESIdtdSizePolicy )
                if cL["UI"] == "TextEdit":
                    cL["EWidget"] = QTextEdit( parent=self )
                    cL["EWidget"].setSizePolicy( ESIdttSizePolicy )
                elif cL["UI"] == "ComboBox":
                    # cL["EWidget"] = QComboBox( parent = self )
                    cL["EWidget"] = itemComboBox( None )  # , parent = self  )
                    cL["EWidget"].setSizePolicy( ESIdtdSizePolicy )
                else:
                    cL["EWidget"] = QLineEdit( parent=self, text="" )
                    cL["EWidget"].setSizePolicy( ESIdtdSizePolicy )
                cL["EWidget"].setVisible( False )
    
                if cL["UI"] == "Entry":
                    cL["EWidget"].textChanged[str].connect( lambda *_, cL=cL : self.entryChanged( *_, cL=cL ) )
                elif cL["UI"] == "ComboBox" :
                    # cL["EWidget"].currentIndexChanged.connect( lambda *_, cL=cL : self.comboChanged( *_, cL=cL ) )
                    cL["EWidget"].activated.connect( lambda *_, cL=cL : self.comboChanged( *_, cL=cL ) )
                    pass
                
                if cL["UI"] != "ComboBox":
                    cL["LWidget"].setAlignment( ESIalignTopLeft )
                    cL["EWidget"].setAlignment( ESIalignTopLeft )
                ipVL.addWidget( iLb, rowNo, 0 )
                ipVL.addWidget( cL["LWidget"], rowNo, 1 )  # Add both static Label and Edit widgets in same location but hide the non-edit one initially
                ipVL.addWidget( cL["EWidget"], rowNo, 1 )
                rowNo += 1
        
        self.iButtons = {}
        for bFunc in [ ["Edit",False,"Edit current item",self.edMode],
                       ["New", True,"Add a new item",self.edNew],
                       ["Save",False,"Save changes and exit from editing",self.edOk],
                       ["Cancel",False,"Abandon changes and exit from editing",self.edCancel]
                       ]:
            iButton = QPushButton(parent=self, text=bFunc[0])
            iButton.setMaximumSize( ESIbuttonMaxSize )
            iButton.clicked.connect( bFunc[3] )
            iButton.setToolTip( bFunc[2] )
            iButton.setEnabled( bFunc[1] )
            ipBL.addWidget( iButton )
            self.iButtons[bFunc[0]] = iButton

        ipRows.addLayout( ipVL )
        ipRows.addLayout( ipBL )
        self.setLayout( ipRows )

    def edMode( self ):
        self.editmode = True
        self.newItem  = False

        self.refresh( (True,False) )
        self.itmLW.setEnabled( False )

    def edNew( self ):
        self.editmode = True
        self.newItem  = True
        # Get value for index of new item       
        self.indexVal = self.dB.nextItemID()
        self.refresh( (True,True) )
        self.itmLW.setEnabled( False )

    def edOk( self ):
        if self.itemEdited:
            
            if not self.newItem:   # UPDATE EXISTING ITEM
                print("Updating item for item %s" % (self.indexVal))
                updateSql = "update ITEMS set "
                firstField = True

                for cL in self.itpFields:
                    if cL["PanelVisible"] == 2:
                        if cL["UI"] == "TextEdit":
                            clModified = cL["EWidget"].document().isModified()
                            newVal = re.sub( "\'", "\'\'", cL["EWidget"].document().toPlainText() )
                        elif cL["UI"] == "ComboBox" :
                            clModified = cL["Edited"]
                            newVal = cL["EWidget"].valueToId( cL["EWidget"].currentText() )
                        else:
                            clModified = cL["EWidget"].isModified()
                            newVal = re.sub( "\'", "\'\'", cL["EWidget"].text() )
                        if clModified:
                            if not firstField:
                                updateSql = updateSql + ", "
                            else:
                                firstField = False
                            print("%s: %s" % (cL["Field"],newVal))
                            updateSql = "%s\"%s\" = \'%s\'" %( updateSql, cL["Field"], newVal)
                        cL["Edited"] = False
                updateSql = "%s where %s = \'%s\'" % ( updateSql, self.itpKey["Field"], self.indexVal )
               
                if firstField:
                    self.txtQW.insert( "Nothing to update")
                else:
                    self.txtQW.insert( "Update: %s" %( updateSql))
                    try:
                        self.dB.cur.execute( updateSql ) 
                        self.dB.con.commit()
                    except:
                        self.txtQW.insert( "ERROR ON DB UPDATE." )

            else:   # CREATING NEW ITEM
                print("Inserting item %s"  % (self.indexVal))
                insertSql = "INSERT INTO ITEMS (%s" % (self.itpKey["Field"])
                insertVals= " VALUES (\'%s\'" % (self.indexVal)
                for cL in self.itpFields:
                    newVal = ""
                    if cL["PanelVisible"] == 2:
                        if cL["UI"] == "TextEdit":
                            newVal = re.sub( "\'", "\'\'", cL["EWidget"].document().toPlainText() )
                        elif cL["UI"] == "ComboBox":
                            newVal = cL["EWidget"].valueToId( cL["Value"] )
                        else:
                            newVal = re.sub( "\'", "\'\'", cL["EWidget"].text() )
                    elif not cL["Default"] == "":   # Assign default values to hidden fields on new items
                        newVal = re.sub( "\'", "\'\'", cL["Default"] )
                    
                    if not newVal is None:
                        if not newVal == "":
                            insertSql  = "%s,\"%s\"" % (insertSql, cL["Field"] )
                            insertVals = "%s,\'%s\'" % ( insertVals, newVal )
                        
                    
                insertSql  = ( "%s, %s" % (insertSql, "Category" ) )
                insertVals = ( "%s, \'%s\'" % (insertVals, self.dB.category ) )
                insertSql = insertSql + ") " + insertVals + ")"
                self.txtQW.insert( "Insert: %s" %( insertSql))
                try:
                    self.dB.cur.execute( insertSql ) 
                    self.dB.con.commit()
                except:
                    self.txtQW.insert( "ERROR ON DB INSERT." )
                
            # Update row in the tree view - currently just reloads the whole dataset from the db
                
            self.dB.refresh()
            self.itmLW.refreshModel()

            # self.itmLW.model.reloadItems()
                                
        self.editmode = False
        print("Resetting buttons in item edit ok")
        self.refresh( (False,False) )
        self.itmLW.setEnabled( True )
        self.itemEdited = False
        
    def edCancel( self ):
        self.editmode = False
        self.newItem  = False

        self.refresh( (False,False) )
        self.itmLW.setEnabled( True )
        self.itemEdited = False
               
    def registerList( self, itmLW ):
        self.itmLW = itmLW
        sect = 1
        for cL in self.itpFields:
            if cL["UI"] == "ComboBox":
                # print ("Looking up sect %d list for %s" % (sect, cL["Field"]))
                self.lookup = self.itmLW.model.columnLookup( sect )
                cL["EWidget"].addLookup( self.lookup )               
            sect += 1
              
    def setID( self, keyValue ):    # Set ID of currently selected item
        self.indexVal = keyValue

    def comboChanged( self, *args, cL ):
        pickedIndex = args[0]
        # On first call there won't be a lookup as yet
        if not cL["EWidget"].lookup is None:
            lookupVal = cL["EWidget"].lookup.loc[pickedIndex,'Label']        
            print( "cL[%s] changed to %s" %(cL["Field"], lookupVal) )
        else:
            lookupVal = ""
        cL["Edited"] = True
        cL["Value"]  = lookupVal
        self.itemEdited = True
        
    def entryChanged( self, *args, cL ): # callback for changes in lineedit

        eText = args[0]
        # print("change in %s" %( eText ))
        self.itemEdited = True
        self.resetOK()

        if len(eText) > 0:
            self.itemEdited = True
            cL["Edited"] = True
            itemStateChange = False
            if cL["Type"] == "int":
                try:
                    testI = int( eText )
                    # Apply data specific range test (temporary hard code)
                    # sets invalid state whilst entry is being hovered but resets
                    # on leaving field - needs to be persistent, may have to look
                    # at how hover state resets valid
                    # NOTE: We aren't using QIntValidator and other built in validators
                    # at the moment, those might be smarter?
                    
                    if not cL["Validator"] is None :
                        if cL["Validator"][0] == "Range":
                            rMin = int(cL["Validator"][1][0])
                            rMax = int(cL["Validator"][1][1])
                            if testI < rMin or testI > rMax:
                                # cL["Widget"].state( ['invalid'] ) TK concept could be used to change widget appearance to flag error visually
                                self.txtQW.insert( "%s (%d) not in range %d to %d" % ( cL["Field"], testI, int(rMin), int(rMax) ) )
                                if cL["Valid"]:
                                    itemStateChange = True
                                    cL["Valid"] = False    # keep separate track of validity as ttk widget state is a bit unpredictable
                            else:
                                if not cL["Valid"]:
                                    itemStateChange = True
                                    cL["Valid"] = True    # keep separate track of validity as ttk widget state is a bit unpredictable
                except (ValueError, TypeError ):
                    self.txtQW.insert( "Invalid integer number (%s)"  % eText )
                    itemStateChange = True
                    cL["Valid"] = False    # keep separate track of validity as ttk widget state is a bit unpredictable
                    # cL["Widget"].config( validate = "all")
                    
            elif cL["Type"] == "float":
                try:
                    testF = float( eText )
                    if not cL["Validator"] is None :
                        if cL["Validator"][0] == "Range":
                            rMin = float(cL["Validator"][1][0])
                            rMax = float(cL["Validator"][1][1])
                            if testF < rMin or testF > rMax:
                                # cL["Widget"].state( ['invalid'] ) TK concept could be used to change widget appearance to flag error visually
                                self.txtQW.insert( "%s (%d) not in range %f to %f" % ( cL["Field"], testF, float(rMin), float(rMax) ) )
                                if cL["Valid"]:
                                    itemStateChange = True
                                    cL["Valid"] = False    # keep separate track of validity as ttk widget state is a bit unpredictable
                            else:
                                if not cL["Valid"]:
                                    itemStateChange = True
                                    cL["Valid"] = True    # keep separate track of validity as ttk widget state is a bit unpredictable
                                    self.txtQW.insert( "%s (%d) now correct in range %f to %f" % ( cL["Field"], testF, float(rMin), float(rMax) ) )
                except (ValueError, TypeError ):
                    self.txtQW.insert( "Invalid floating point number (%s)"  % eText )
                    itemStateChange = True
                    cL["Valid"] = False    # keep separate track of validity as ttk widget state is a bit unpredictable
                    # cL["Widget"].config( validate = "all")
    
            if itemStateChange:
                self.resetOK( )
            
    def resetOK( self ): # Reset status of OK/Save button based on current item validity
    
        allValid = True
        for cL in self.itpFields:
            if not cL["Valid"]:
                allValid = False
                break
        if allValid:
           self.iButtons["Save"].setEnabled( True)
        else:
           self.iButtons["Save"].setEnabled( False )

    def refresh( self, itmArgs ): # ItemPanel refresh - gets data directly from dB dataframe, not the table model
        setEdMod = itmArgs[0]
        isNew    = itmArgs[1]
        self.itpKey["Widget"].setText( str(self.indexVal) )
        if not isNew: # This is failing on the str indexVal entries - more likely a pandas than a qt problem.       
            itpRow = self.dB.itmDf[self.dB.itmDf[self.itpKey["Field"]] == self.indexVal ]
        else:
            itpRow = self.dB.defaultDf()
        rowNo  = 1
        self.editMode = setEdMod
        for cL in self.itpFields:
            if not cL["PanelVisible"] == 0:
                curValue = str(itpRow[cL["Field"]].values[0])
                if self.editMode and cL["PanelVisible"] == 2: # This refreshes directly from the pd, not using the tablemodel translation
                    cL["LWidget"].setVisible( False )
                    cL["Edited"] = False if not isNew else True
    
                    if not cL["EWidget"] is None:
                        cL["EWidget"].setVisible( True )
                        if cL["UI"] == "TextEdit":  # TextEdit could also be PlainTextEdit requiring setPlainText ...
                            cL["EWidget"].setText( curValue )
                        elif cL["UI"] == "ComboBox":  # This just accumulates items, actually just need to set index to current value)
                            if curValue == "":
                                curValue = cL["Default"]
                            curValue = cL["EWidget"].idToValue( curValue )
                            if not curValue is None:
                                cL["EWidget"].setCurrentText( curValue )
                        else:
                            cL["EWidget"].setText( curValue )
                    else:
                        cL["dVar"] = curValue
                        cL["EWidget"] = QLineEdit( text= cL["dVar"] ) 
                else:
                    if cL["UI"] == "ComboBox":
                        curValue = cL["EWidget"].idToValue( curValue )
                    cL["LWidget"].setText( curValue )
                    if not cL["EWidget"] is None:
                        cL["EWidget"].setVisible( False )
                    cL["LWidget"].setVisible( True )
                rowNo += 1
        if setEdMod:
            self.iButtons["Edit"].setEnabled( False )
            self.iButtons["New"].setEnabled( False )
            self.iButtons["Save"].setEnabled( True )
            self.iButtons["Cancel"].setEnabled( True )
        else:
            self.iButtons["Edit"].setEnabled( True )
            self.iButtons["New"].setEnabled( True )
            self.iButtons["Save"].setEnabled( False)
            self.iButtons["Cancel"].setEnabled( False )
            
class itemTableModel( QAbstractTableModel ):
    def __init__(self, dB=None, itmEW=None, imgList=None):

        QAbstractTableModel.__init__(self)
        self.dB       = dB
        self.itpKey   = dB.itpKey
        self.itemCols = dB.itpFields
        self.itmEW    = itmEW
        self.imgList  = imgList
        self.keyValue = None
        self.tableWidth = 0
        tCols = [ self.itpKey["Field"] ]
        colWidth = ( self.itemCols[0]["Width"] * ESIcharacterWidth ) + ESIlistColPadding
        # self.setColumnWidth( 0, colWidth )
        self.tableWidth += int(colWidth)
        colNo = 1        
        for cL in self.itemCols:
            tCols.append( cL["Label"] )
            colWidth = ( cL["Width"] * ESIcharacterWidth ) + ESIlistColPadding
            if not cL["ListVisible"] == 0:
                self.tableWidth += int(colWidth)
            # self.setColumnWidth( colNo, colWidth ) 
            # self.setHeaderData(colNo, Qt.Orientation.Horizontal, cL["Label"] )
            colNo += 1
        # self.reloadItems()
            
    def rowCount(self, parent=QModelIndex()):
        return len(self.dB.itmDf)
    
    def columnCount(self, parent=QModelIndex()):
        return 1+len(self.itemCols)
    
    def headerData(self, section, orientation, role):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            # hdrs = [self.itpKey["Field"]]
            # for iCol in self.itemCols:
            #    hdrs.append( iCol["Field"])
            if section == 0 :
                hdr = self.itpKey["Label"]
                # hdr = QLabel( text="ID Test")
            else:
                hdr = self.itemCols[section-1]["Label"]
            return ( hdr )
        else:
            return f"{section}"

    def columnSelector( self, section ):
        if section > 0 and section <= len( self.itemCols ):
            cL = self.itemCols[section-1]
            cSel = { k: cL.get(k, "") for k in ('UI', 'Type', 'ListVisible', 'Validator', 'Validator1', 'Validator2') }
        else:
            cSel = { k: "" for k in ( 'UI', 'Type', 'ListVisible', 'Validator', 'Validator1', 'Validator2' ) }
        return cSel
    
    def columnLookup( self, section ): # Return the lookups dataframe for a column
        lookup = None
        # Section numbers are field numbering +1 as left most column is defined by itpKey
        if section > 0 and section <= len( self.itemCols ):
            cL = self.itemCols[section-1]
            if cL['Validator'][0] == 'Lookup' :
                lookupName = cL['Validator'][1]
                if lookupName == "":
                    qWin.txtQW.insert("ERROR IN DATA - missing lookup in CategoryFields for field %s" % ( cL['Field'] ))
                    print("ERROR IN DATA - missing lookup in CategoryFields for field %s" % ( cL['Field'] ))
                    sys.exit( 1 )
                else:
                    lookup = self.dB.lookups[lookupName]
        return lookup

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        column = index.column()
        row = index.row()
        if role == Qt.ItemDataRole.DisplayRole:
            if column == 0 :
                return( str(self.dB.itmDf.loc[row,self.itpKey["Field"]]))
            else:
                iCol = self.itemCols[column-1]
                if iCol["Type"] == "str" or iCol["Format"] == "":
                    dbValue = str(self.dB.itmDf.loc[row,iCol["Field"]])
                    if not dbValue is None:
                        if not dbValue == "":
                            if not iCol["Validator"] is None:
                                if iCol["Validator"][0] == 'Lookup' :
                                    lookup = self.dB.lookups[iCol['Validator'][1]]
                                    ll = (lookup[lookup['ID'] == dbValue]['Label'])
                                    if not ll.empty :                           
                                        dbValue = ll.item()
                                    else:
                                        qWin.txtQW.insert("ERROR IN DATA - value %s not in lookup %s" % ( dbValue, iCol['Validator'][1] ))
                    else:
                        dbValue = ""
                    return dbValue
                elif iCol["Type"] == "float":
                    return( iCol["Format"] % float(self.dB.itmDf.loc[row,iCol["Field"]] ) )
                elif iCol["Type"] == "int":
                    return( iCol["Format"] % int(self.dB.itmDf.loc[row,iCol["Field"]] ) )
                else:
                    return( "ERROR IN DATA MODEL FORMAT")
                
        elif role == Qt.ItemDataRole.BackgroundRole:
            return QColor( "#ffffff")
        elif role == Qt.ItemDataRole.TextAlignmentRole:
            return Qt.AlignmentFlag.AlignLeft  # Right
    
        return None

class itemList( QWidget ):
    
    def __init__( self, dB, itmEW, imgList ):
        QWidget.__init__(self)

        self.tView = QTableView()
        self.hHeader = self.tView.horizontalHeader()      
        self.vHeader = self.tView.verticalHeader()
        self.hHeader.setSectionResizeMode( QHeaderView.ResizeMode.ResizeToContents )
        self.vHeader.setSectionResizeMode( QHeaderView.ResizeMode.ResizeToContents )
        self.vHeader.setVisible( False )
        self.hHeader.setStretchLastSection(True)
        """
        self.tView.horizontalHeader(). is a QHeaderView https://doc.qt.io/qt-6/qheaderview.html 
        or https://doc.qt.io/qtforpython-5/PySide2/QtWidgets/QHeaderView.html that has hHeader.count() sections.
        The sections are defined in the itemTableModel.headerData()
        """
        
        self.itmEW   = itmEW
        self.dB      = dB
        self.imgList = imgList
        self.refreshModel( )
        self.tableWidth = self.model.tableWidth
        self.filterUis = []
        # QWidget Layout       
        self.main_layout = QVBoxLayout()
        
        itemBox = QGroupBox( "Items" )
        itemBox.setStyleSheet( ESIupperCaptionStyle )
        itemBox.setSizePolicy( ESIitmListSizePolicy )
        itemBoxRows = QGridLayout()
        filterNo = 0
        for sect in range( 0, self.hHeader.count() ):
            cSel = self.model.columnSelector( sect )
            if cSel["ListVisible"] == 2:

                if cSel["UI"] == 'Entry':
                    if cSel["Type"] == 'float':
                        filterUi = itemRangeFilter( dB, (sect - 1) )
                    else:
                        filterUi = itemTextFilter( dB, (sect - 1) )

                elif cSel["UI"] == 'ComboBox' and cSel["Validator"] != "":
                    filterUi = itemComboMultiBox( self.model.columnLookup( sect ), self.dB, (sect-1) )
                else:
                    filterUi = None
                self.filterUis.append( filterUi )
                if not filterUi is None:
                    filterHdr = boxLabel( parent = self, text=self.model.headerData( sect, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole ))
                    itemBoxRows.addWidget( filterHdr, 0, filterNo )
                    itemBoxRows.addWidget( filterUi, 1, filterNo )
                    filterNo += 1

            elif cSel["ListVisible"] == 0:
                self.tView.setColumnHidden( sect, True )
                self.hHeader.hideSection( sect )
        itemBoxRows.addWidget( self.tView, 2, 0, 1, filterNo )
        itemBox.setLayout( itemBoxRows )
        self.tView.setSizePolicy( ESIitmListSizePolicy )
        
        # self.main_layout.addWidget(self.tView.searchBar)
        # self.main_layout.addLayout( itemBoxRows )
        self.main_layout.addWidget(itemBox)
        # self.main_layout.addWidget(self.tView)
        self.tView.clicked.connect( self.pickItem )
        self.setLayout(self.main_layout)
        
    def refreshModel( self ):
        self.model   = itemTableModel( self.dB, self.itmEW, self.imgList )
        self.proxyModel = QSortFilterProxyModel()
        self.proxyModel.setFilterKeyColumn( -1 )
        self.proxyModel.setSourceModel( self.model )
        self.proxyModel.setFilterCaseSensitivity( Qt.CaseSensitivity.CaseInsensitive )
        self.proxyModel.sort( 0, Qt.SortOrder.AscendingOrder ) # Default sort on list startup
        self.tView.setSortingEnabled( True )
        # self.tView.setModel( self.model )

        self.tView.setModel( self.proxyModel )
        # self.tView.searchBar = QLineEdit()
        # self.tView.searchBar.textChanged.connect(self.proxyModel.setFilterFixedString)
        
    def pickItem( self, pickedItem ):
        keyValue = pickedItem.model().index(pickedItem.row(),0).data()
        # self.itmEW.setID( QTableWidgetItem(str(keyValue)) )
        self.itmEW.setID( keyValue ) # has already been forced to pandas dtype string int(keyValue) )
        self.itmEW.refresh( (False,False) )
        self.imgList.loadImages( keyValue, self.getImages( keyValue, False ) )
        
    def getImages( self, itemId, primaryOnly ):
        ImagesCols = [ "AlbumID", "ImageID","ItemID","Rank","Zoomed","Edited","DateTaken","Camera","Background","Rotation" ]

        imgSql = "Select AlbumID,ImageID,ItemID,\"Rank\",Rotation from IMAGES where Rank NOT IN (\"D\") and ItemID = \"%s\"" %(itemId)
        res = self.dB.cur.execute( imgSql )
        images = res.fetchall()
        imgs = [] # ( image file name, rotation flag, availability flag, pixmap, path, rank flag )
        # Should actually look up column dynamically, not rely on being [2]
        # Also need to use Albums table to get prefix
        row = 0
        for iRow in images:
            imgRank = iRow[3]
            rotation   = iRow[4]
            if rotation is None:
                rotation = 0.0
            imgName = iRow[1]
            imgFile = imgName + ".jpg"
            imgPath = self.dB.albumPath + imgFile
            imgRank = iRow[3]
            imgExists = True
            if not primaryOnly or imgRank == "P":
                if not os.path.exists( imgPath ):
                    imgExists = False
                    imgPix = None
                else:
                    imgPix = QPixmap(imgPath)
                    if rotation != 0.0:
                        imgPix = imgPix.transformed(QTransform().rotate(rotation))

            # imgs.append( (imgName,rotation,imgExists,imgPix,imgPath,imgRank) )
                imgs.append({"ImageID":     imgName,
                             "Item":        itemId,
                             "Row":         row,
                             "Rotation":    rotation,
                             "Exists":      imgExists,
                             "Pixmap":      imgPix,
                             "Path":        imgPath,
                             "Rank":        imgRank
                              })
            row += 1
        # self.imgList.loadImages( itemId, imgs )
        return( imgs )


    def exportPPTX( self ):
        pptxFile = "my.pptx"
        print("Exporting current item list to PPTX")
        self.nRows = self.proxyModel.rowCount()
        self.progressWin = QProgressDialog( "Reporting %d items to Powerpoint file %s" % (self.nRows, pptxFile), "Cancel", 0, (self.nRows-1), self )
        self.progressWin.setWindowModality( Qt.WindowModality.WindowModal )
        mHdrs = []
        for col in range( self.proxyModel.columnCount()):
            mHdrs.append( self.proxyModel.headerData( col, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole ) )

        pptx = pptxCatalog( pptxFile, None )
        ( pImgWid, pImgHt ) = pptx.imageSize()
        pptxImageSizeDpi = QSize( int(pImgWid * ESIimgPptxDpi), int(pImgHt * ESIimgPptxDpi)  )
        
        for row in range( self.nRows ):
            itmData = {}
            for col in range( self.proxyModel.columnCount()):
                index = self.proxyModel.index( row, col )
                itmData[mHdrs[col]] = self.proxyModel.data( index )
            imgData = self.getImages( itmData["ID"], True )
            if len(imgData) > 0:
                print("Adding image %s for item %s" % ( imgData[0]["ImageID"], itmData["ID"]))
                tempPix = imgData[0]["Pixmap"].scaled(
                    pptxImageSizeDpi,
                    aspectRatioMode = QtCore.Qt.AspectRatioMode.KeepAspectRatio)
                tempPix.save( "temp_pixmap.jpg", "JPG")
                imgSizeRatio = tempPix.size().height()/tempPix.size().width()
                pptx.addItemWithImage( itmData, imgData[0]["ImageID"], "temp_pixmap.jpg", imgSizeRatio )
            else:
                print("Error in getting image for item %s" % (itmData["ID"]))
            self.progressWin.setValue( row )
            if self.progressWin.wasCanceled():
                break
        pptx.saveOutput()


# Text box for string item filters

class itemTextFilter( QLineEdit ):

    
    def __init__(self, dB, dbSection ):
        super().__init__()
        self.dB = dB
        self.dbSection = dbSection
        self.textChanged[str].connect( self.filterChanged )
        self.typeCounter = 0
        self.editingFinished.connect( self.finished )

    def finished( self, *args ):
        self.typeCounter = 0
        editedText = self.newText.strip()
        filters = [ editedText ]
        self.dB.columnFilters( self.dbSection, 'Text', filters )
        self.setText( editedText )
        self.dB.refresh()
        self.parent().parent().refreshModel()
        
    def filterChanged( self, *args ):
        self.newText = args[0]
        """
        if self.typeCounter == 0:
            self.rTimer = QtCore.QTimer.singleShot( ESItimerFilterText, self.filterTimer )
            self.typeCounter = 1
        elif self.typeCounter == 1:
            self.typeCounter = 2  # This tells the timer event that there have been more type events since being triggered
        # QtWidgets.QMainWindow.resizeEvent( self, event )                  
        """
    """    
    def filterTimer( self ):  
        if self.typeCounter == 1:
            # Typing deemed complete
            self.typeCounter = 0
            editedText = self.newText.strip()
            filters = [ editedText ]
            self.dB.columnFilters( self.dbSection, 'Text', filters )
            self.setText( editedText )
            self.dB.refresh()
            self.parent().parent().refreshModel()
        else:
            self.rTimer = QtCore.QTimer.singleShot( ESItimerFilterText, self.filterTimer )
            self.typeCounter = 1
    """
    
# Text box for float item filters - 
# Does not refresh SQL on certain lose focus from linedit to main list screen,
# maybe as the focusTimer is not working 100% correctly... maybe you don't need
# the focusTimer if we use editingFinished signal! 


class itemRangeFilter( QWidget ):
    def __init__( self, dB, dbSection ):
        QWidget.__init__(self)
        self.dB = dB
        self.dbSection = dbSection
        self.fromFilter = itemRangeFloatFilter( "From" )
        toLabel = QLabel( "to" )
        self.toFilter   = itemRangeFloatFilter( "To" )
        rangeLayout     = QHBoxLayout()
        rangeLayout.addWidget( self.fromFilter )
        rangeLayout.addWidget( toLabel )
        rangeLayout.addWidget( self.toFilter )
        self.setLayout( rangeLayout )
        self.inFrom = False
        self.inTo   = False
        self.fromValue = 0.0
        self.toValue   = 0.0
        self.inCounter = 0
        self.toBlank   = True
        self.fromBlank = True
        
        # self.focusOutEvent.connect(partial(self.rangeChanged)) # WrONG
        
    def focusInEvent(self, rangeEnd, event):
        super().focusInEvent(event)
        if self.inCounter == 0:
            self.rTimer = QtCore.QTimer.singleShot( ESItimerFilterText, self.focusInTimer )
            self.inCounter = 1
        elif self.inCounter == 1:
            self.inCounter = 2  # This tells the timer event that there have been more type events since being triggered

        if rangeEnd == "To":
            self.inTo = True
        else:
            self.inFrom = True

    def focusInTimer( self ):  
        if self.inCounter == 1:
            # Typing deemed complete - have we exited from both from and to entries?
            if not self.inTo and not self.inFrom:
                self.inCounter = 0
                if self.fromBlank or self.toBlank:
                    filters = [ None ]
                else:
                    filters = [ self.fromValue, self.toValue ]
                self.dB.columnFilters( self.dbSection, 'Float', filters )
                self.dB.refresh()
                self.parent().parent().refreshModel()
            else:
                self.rTimer = QtCore.QTimer.singleShot( ESItimerFilterText, self.focusInTimer )
                self.inCounter = 1
        else:
            self.rTimer = QtCore.QTimer.singleShot( ESItimerFilterText, self.focusInTimer )
            self.inCounter = 1

    def focusOutEvent(self, rangeEnd, newValue, isBlank, event):
        if not event is None:
            super().focusOutEvent(event)
        if rangeEnd == "To":
            self.inTo = False
            self.toValue = newValue
            self.toBlank = isBlank
        else:
            self.inFrom = False
            self.fromValue = newValue
            self.fromBlank = isBlank
        self.inCounter = 0
            
class itemRangeFloatFilter( QLineEdit ):

    def __init__(self, rangeEnd ):
        super().__init__()
        self.oldText = ""
        self.newFloat= 0.0
        self.typeCounter = 0
        self.isBlank = True
        self.rangeEnd = rangeEnd
        self.textChanged[str].connect( self.filterChanged )
        self.editingFinished.connect( self.finished )
        self.setStyleSheet( ESIinputFieldOkStyle )

    def finished( self, *args ):
        self.parent().focusOutEvent( self.rangeEnd, self.newFloat, self.isBlank, None )
        
    def filterChanged( self, *args ):
        # Reject entries that aren't valid floating point numbers
        
        try:
            testStr = args[0].strip() 
            if not testStr == "":
                testF = float( testStr )
                self.newFloat = testF
                self.isBlank    = False
                self.oldText  = testStr
            else:
            
                self.newFloat = 0.0
                self.isBlank    = True
                self.oldText  = ""
            # Should set a field ok style here
            self.setStyleSheet( ESIinputFieldOkStyle )
        except (ValueError, TypeError ):
            # Should set a field alert style here
            self.setText( self.oldText )
            self.setStyleSheet( ESIinputFieldWarnStyle )
            styleTimer = QtCore.QTimer.singleShot( ESItimerWarnStyle, self.resetStyle )

    def resetStyle( self ):
        self.setStyleSheet( ESIinputFieldOkStyle )
        
    def focusInEvent(self, event):
        super().focusInEvent(event)
        self.parent().focusInEvent( self.rangeEnd, event )

    def focusOutEvent(self, event):
        super().focusOutEvent(event)
        self.parent().focusOutEvent( self.rangeEnd, self.newFloat, self.isBlank, event )

# Checkable Combobox (from stackexchange 350148 )

class CheckableComboBox(QComboBox):

    # Subclass Delegate to increase item height
    class Delegate(QStyledItemDelegate):
        def sizeHint(self, option, index):
            size = super().sizeHint(option, index)
            size.setHeight(20)
            return size

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Make the combo editable to set a custom text, but readonly
        self.setEditable(True)
        self.lineEdit().setReadOnly(True)
        # Make the lineedit the same color as QPushButton
        palette = qtApp.palette()
    #    palette.setBrush(QPalette.Base, palette.button())
      
        self.lineEdit().setPalette(palette)

        # Use custom delegate
        self.setItemDelegate(CheckableComboBox.Delegate())

        # Update the text when an item is toggled
        self.model().dataChanged.connect(partial(self.updateText,False))

        # Hide and show popup when clicking the line edit
        self.lineEdit().installEventFilter(self)
        self.closeOnLineEditClick = False

        # Prevent popup from closing when clicking on an item
        self.view().viewport().installEventFilter(self)

    def resizeEvent(self, event):
        # Recompute text to elide as needed
        self.updateText(False)
        super().resizeEvent(event)

    def eventFilter(self, object, event):

        if object == self.lineEdit():
            if event.type() == QEvent.Type.MouseButtonRelease:
                if self.closeOnLineEditClick:
                    self.hidePopup()
                else:
                    self.showPopup()
                return True
            return False

        if object == self.view().viewport():
            if event.type() == QEvent.Type.MouseButtonRelease:
                index = self.view().indexAt(event.pos())
                item = self.model().item(index.row())

                if item.checkState() == Qt.CheckState.Checked:
                    item.setCheckState(Qt.CheckState.Unchecked)
                else:
                    item.setCheckState(Qt.CheckState.Checked)
                return True
        return False

    def showPopup(self):
        super().showPopup()
        # When the popup is displayed, a click on the lineedit should close it
        self.closeOnLineEditClick = True

    def hidePopup(self):
        super().hidePopup()
        # Used to prevent immediate reopening when clicking on the lineEdit
        self.startTimer(100)
        # Refresh the display text when closing
        self.updateText( True )

    def timerEvent(self, event):
        # After timeout, kill timer, and reenable click on line edit
        self.killTimer(event.timerId())
        self.closeOnLineEditClick = False

    def updateText(self, updateFilters):
        texts = []
        filters = []
        for i in range(self.model().rowCount()):
            if self.model().item(i).checkState() == Qt.CheckState.Checked:
                texts.append(self.model().item(i).text())
                filters.append(self.valueToId(self.model().item(i).text()))
        text = ", ".join(texts)

        if updateFilters:
            self.dB.columnFilters( self.dbSection, 'List', filters )
            self.dB.refresh()
            self.parent().parent().refreshModel()

        # Compute elided text (with "...")
        metrics = QFontMetrics(self.lineEdit().font())
        elidedText = metrics.elidedText(text, Qt.TextElideMode.ElideRight, self.lineEdit().width())
        self.lineEdit().setText(elidedText)

    def addItem(self, text, data=None):
        item = QStandardItem()
        item.setText(text)
        if data is None:
            item.setData(text)
        else:
            item.setData(data)
        item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
        item.setData(Qt.CheckState.Unchecked, Qt.ItemDataRole.CheckStateRole)
        self.model().appendRow(item)

    def addItems(self, texts, datalist=None):
        for i, text in enumerate(texts):
            try:
                data = datalist[i]
            except (TypeError, IndexError):
                data = None
            self.addItem(text, data)

    def currentData(self):
        # Return the list of selected items data
        res = []
        for i in range(self.model().rowCount()):
            if self.model().item(i).checkState() == Qt.CheckState.Checked:
                res.append(self.model().item(i).data())
        return res       


# Combobox needs to look up label from code, present box of labels and return selected code only

class itemComboMultiBox( CheckableComboBox ): # QComboBox ):
    
    def __init__( self, itemLookup, dB, dbSection ):
        super().__init__()
        if not itemLookup is None:
            self.addLookup( itemLookup )
        self.dB = dB
        self.dbSection = dbSection

    def addLookup( self, itemLookup ):
        self.lookup = itemLookup
        listItems = itemLookup['Label'].tolist()
        self.addItems( listItems )
        
    def valueToId( self, value ):
        labelId = ""
        if not value is None:
            if not value == "":
                ll = (self.lookup[self.lookup['Label'] == value]['ID'])
                if not ll.empty :                           
                    labelId = ll.item()
                else:
                    qWin.txtQW.insert("ERROR IN DATA - value %s not a label in lookup" % ( value ))
            
        return labelId

    def idToValue( self, valId ):
        idLabel = ""
        if not valId is None:
            if not valId == "":
                ll = (self.lookup[self.lookup['ID'] == valId]['Label'])
                if not ll.empty :                           
                    idLabel = ll.item()
                else:
                    qWin.txtQW.insert("ERROR IN DATA - value %s not an ID in lookup" % ( valId ))
        return idLabel

class itemComboBox( QComboBox ): # QComboBox ):
    
    def __init__( self, itemLookup ):
        super().__init__()
        if not itemLookup is None:
            self.addLookup( itemLookup )

    def addLookup( self, itemLookup ):
        self.lookup = itemLookup
        listItems = itemLookup['Label'].tolist()
        self.addItems( listItems )
        
    def valueToId( self, value ):
        labelId = ""
        if not value is None:
            if not value == "":
                ll = (self.lookup[self.lookup['Label'] == value]['ID'])
                if not ll.empty :                           
                    labelId = ll.item()
                else:
                    qWin.txtQW.insert("ERROR IN DATA - value %s not a label in lookup" % ( value ))
        return labelId

    def idToValue( self, valId ):
        idLabel = ""
        if not valId is None:
            if not valId == "":
                ll = (self.lookup[self.lookup['ID'] == valId]['Label'])
                if not ll.empty :                           
                    idLabel = ll.item()
                else:
                    qWin.txtQW.insert("ERROR IN DATA - value %s not an ID in lookup" % ( valId ))
        return idLabel
    
class boxLabel( QLabel ): # Header for a box or other subsection

    def __init__( self,  *args, **kwargs):
        
        super().__init__(*args, **kwargs)
        # self.setMaximumSize( ESIbuttonMaxSize )
        self.setAlignment( ESIalignBottomLeft )
        self.setStyleSheet( ESIupperCaptionStyle )

class scrollText( QTextEdit ):   # Scrolling text output window
    
    def __init__( self, parent ):
        super().__init__( parent )
        self.setMinimumSize( ESItextWinMinSize )
        self.setStyleSheet( "color:%s; background-color: white" % ( ESIcaptionFontColor ))
        self.setStyleSheet( ESIcaptionFont )
        # self.setSizePolicy( ESIdtSizePolicy )
        self.setText("Welcome.")
        
    def insert( self, text ):
        self.append( text )
        self.moveCursor( QtGui.QTextCursor.MoveOperation.End )
        self.ensureCursorVisible()
        
class infoWindow( QWidget ):   # Free floating information window
    
    def __init__( self, winTitle, winText ):

        super().__init__()
        self.setWindowTitle( "eStoreInventory - " + winTitle )
        iLayout = QVBoxLayout()
        self.iLabel  = QLabel( winText )
        iLayout.addWidget( self.iLabel )
        self.setLayout( iLayout )
        
    def addText( self, text ):
        lText = self.iLabel.text() + "\n" + text
        self.iLabel.setText( lText )
        

#---------------------------------------------------------------------------------
# Main application class

class uiLayout( QMainWindow ):
    
    def __init__( self, baseDir, qtApp ):

        # QT Window initialize
        
        super().__init__()
        self.baseDir = baseDir
        self.qtApp   = qtApp
        self.resizeCounter = 0
        self.setWindowTitle( "eStore Inventory")
        category = "Beauties"
        # category = "Books"
        # category = "Vinyl"
        dtlMaxWidth = ESIdataGroupMaxSize.width()               
        # Layout options include QHBox QVBox and QGridLayout()
        
        apL = QHBoxLayout()   # Data vs images frames
        dtL = QVBoxLayout()   # Data sub frames
        imL = QHBoxLayout()   # Image sub frames
        appBL = QHBoxLayout() # Applications control button list
        
# Scrolling monitor text window

        monBox = QGroupBox()
        monBox.setStyleSheet( ESIupperCaptionStyle )
        monL = QVBoxLayout()
        txtHdr = boxLabel( parent = self, text = "Monitor" )
        self.txtQW = scrollText( parent=monBox )
        monL.addWidget( txtHdr)
        monL.addWidget( self.txtQW )
        monBox.setLayout( monL )
        
        # Database access window
        
        homePath = "/home/volker/Dropbox/private_python/" if os.name == 'posix' else "c:\\Users\\volke\\DropBox\\private_python\\"
        self.dB = itemDb( "%s%s" % (homePath,"estore.db"), category, self.txtQW ) # "Vinyl" "Beauties"
        # Image management window
        
        self.imgView  = imageView( self.dB, parent=self )
        self.imgList  = imageList( self.imgView, parent=self )

        self.imgView.setMaximumSize( ESIimageMaxSize )
        
        self.itmEW = itemPanel( self.dB, parent=self )
        
        self.itmLW = itemList( self.dB, self.itmEW, self.imgList ) # parent=self  )

        self.itmEW.registerList( self.itmLW ) # not self.itemList.tV )

        
        if self.itmLW.tableWidth < dtlMaxWidth:
            ESIdataGroupMaxSize.setWidth( self.itmLW.tableWidth )
                
        self.txtQW.setMaximumSize( ESIdataGroupMaxSize ) 
        self.itmEW.setMaximumSize( ESIdataGroupMaxSize ) 
        self.itmLW.setMaximumSize( ESIdataGroupMaxSize )
        monBox.setMaximumSize( ESIdataGroupMaxSize )
        self.itmEW.setSizePolicy( ESIdtSizePolicy )
        self.itmLW.setSizePolicy( ESIitmListSizePolicy )
        monBox.setSizePolicy( ESIdtSizePolicy )

        """
        exitButton = QPushButton(parent=self, text="Quit")
        exitButton.setMaximumSize( ESIbuttonMaxSize )
        exitButton.clicked.connect( self.appQuit )
        """
        appBL.setAlignment( QtCore.Qt.AlignmentFlag.AlignLeft )

        for w in [ # imL,
                   self.imgView, self.imgView.picImage, self.imgView.picInfo ]:
            w.setSizePolicy( ESIimgSizePolicy )
            
        # exitButton.setSizePolicy( ESIimlSizePolicy ) # To get fixed horizontal
        self.imgList.setSizePolicy( ESIimlSizePolicy )
 
        dtL.addWidget( self.itmEW )
        dtL.addWidget( self.itmLW )
        dtL.addWidget( monBox)
        dtL.addLayout( appBL )

        imL.addWidget( self.imgList )
        imL.addWidget( self.imgView )
        apL.addLayout( dtL )
        apL.addLayout( imL )
        
        # appBL.addWidget( exitButton )
        
        apWidget = QWidget()
        apWidget.setLayout( apL )
        
        # Menubar - should ideally be a floating toolbar on the left hand side of screen only;
        # or maybe make it easier to float the image list + image into another window.

        menuBar = QMenuBar( self )
        fileMenu= QMenu( "&File", self )   # & precedes the hot key for this menu option
        helpMenu= QMenu( "&Help", self )
        
        exportAction = QAction( "E&xport", self )
        exportAction.triggered.connect( self._export )

        prefAction = QAction( "&Preferences", self )
        prefAction.triggered.connect( self._preferences )

        exitAction = QAction( "&Exit", self )
        exitAction.triggered.connect( self.appQuit ) # sys.exit(0)) # self.app.quit() )

        fileMenu.addAction( exportAction )
        fileMenu.addAction( prefAction )
        fileMenu.addAction( exitAction )       

        configReportAction = QAction( "&Configuration report", self )
        configReportAction.triggered.connect( self._reportConfiguration )

        helpAboutAction = QAction( "&About", self )
        helpAboutAction.triggered.connect( self._helpAbout )

        helpMenu.addAction( configReportAction )
        helpMenu.addAction( helpAboutAction )

        menuBar.addMenu( fileMenu )
        menuBar.addMenu( helpMenu )
        self.setMenuBar( menuBar )

        self.setCentralWidget( apWidget )
                       
        return
    
    def _reportConfiguration( self ):
        self.aboutWin = infoWindow( "Configuration", "eStore Configuration\nBase path:\t%s\nCategory:\t%s\nDB path: \t%s\nAlbum path:\t%s" % (self.baseDir, self.dB.category, self.dB.dbPath, self.dB.albumPath ) )
        silentFields = ""
        silentComma  = ""
        self.aboutWin.addText( "Visible fields - ")        
        for cL in self.dB.itpFields:
            if not cL["ListVisible"] + cL["PanelVisible"] == 0:
                filterType = ""
                if not cL["Filters"] is None:
                    if not cL["Filters"][0] == "":
                        filterType = "Filtering on " + cL["Filters"][0]
                self.aboutWin.addText( "%s - [%s %d] %s, filter %s" % (cL["Field"], cL["Type"], cL["Width"], 
                        "Valid" if cL["Valid"] else "Invalid", 
                        filterType ))
            else:
                silentFields = silentComma + silentFields + cL["Field"]
                silentComma  = ", "
        if silentFields != "":
            self.aboutWin.addText( "Silent fields - \n%s" % ( silentFields ) )
        self.aboutWin.show()
               
    def _helpAbout( self ):
        self.aboutWin = infoWindow( "About", "eStoreInventory - Volker Hirsinger 2024")
        self.aboutWin.show()

    def _preferences( self ):
        print( "Preferences" )
    
    def _export( self ):
        print("Exporting to ... ")
        # For a start just export out of the item list as this leverages the
        # current filtering. Later maybe have alternate ways of selecting.
        
        self.itmLW.exportPPTX( )
        
    def resizeEvent( self, event ):
        self.size1 = event.oldSize()
        self.size2 = event.size()
        # self.txtQW.insert("Resize state %d from (%d,%d) to (%d,%d)" %( self.resizeCounter, self.size1.width(), self.size1.height(), self.size2.width(), self.size2.height() ))
        if self.resizeCounter == 0:
            # self.imgView.picImage.clear()
            # self.txtQW.insert( "Resize timer restarted.")  # This can probably be reduced to a single call to singleShot from resizeTimer itself.
            self.rTimer = QtCore.QTimer.singleShot( ESItimerResize, self.resizeTimer )
            self.resizeCounter = 1
        elif self.resizeCounter == 1:
            self.resizeCounter = 2   # This tells the timer event that there have been more resize callbacks since being triggered.
        QtWidgets.QMainWindow.resizeEvent( self, event )

    def resizeTimer( self ):  
        # self.txtQW.insert("call Resize count %d from (%d,%d) to (%d,%d)" %( self.resizeCounter, self.size1.width(), self.size1.height(), self.size2.width(), self.size2.height() ))
        if self.resizeCounter == 1:
            # self.txtQW.insert( "Resize complete - now scale the image.")
            self.imgView.refresh( self.imgView.img )
            self.resizeCounter = 0
        else:
            # self.txtQW.insert( "Resize timer restarted in callback.")
            self.rTimer = QtCore.QTimer.singleShot( ESItimerResize, self.resizeTimer )
            self.resizeCounter = 1
    
    def appQuit( self ):
        self.qtApp.quit()
        self.dB.close()
        # sys.exit( 0 )

    def takePhoto(self):
        # inTxt = gTInStr.get()
        # gTOut.set( inTxt + " Yes")
        picFile = "%s_%d_%d" % (self.dB.ImgPrefix, self.dB.ImgSequence, self.dB.ItmSequence )
        picPath = "%s%s.jpg" % (self.dB.albumPath, picFile )
        gphotoArgs = [ 
            "gphoto2",
            "--capture-image-and-download",
            "--no-keep",
            "--force-overwrite",
            "--filename=" + picPath ]
        clickProc = subprocess.run( gphotoArgs, capture_output=True, text=True )
        if not clickProc.returncode == 0 or not clickProc.stderr == "":
            print("Error code %d on taking photo: %s" %(clickProc.returncode, clickProc.stdout + ' ' + clickProc.stderr))
            gphotoError = re.sub( ".*\*\*\* Error: ", "", clickProc.stderr)
            gphotoError = re.sub( "\**", "", gphotoError )
            qWin.txtQW.insert("Error on taking photo: %s" % (gphotoError))
            self.imgView.refresh( ("NotFound",0))
        else:
            print("Return code %d on taking photo: %s" %(clickProc.returncode, clickProc.stdout + ' ' + clickProc.stderr))
            qWin.txtQW.insert("Return code %d on taking photo: %s" %(clickProc.returncode, clickProc.stdout + ' ' + clickProc.stderr))
    

            self.dB.ImgSequence = self.dB.ImgSequence + 1
            self.dB.ItmSequence = self.dB.ItmSequence + 1
            updateSql = "update CATEGORIES set ItmSequence=%d,ImgSequence=%d where Category = \"%s\"" % ( self.dB.ItmSequence, self.dB.ImgSequence, self.dB.category )
            self.dB.cur.execute( updateSql )
            self.imgView.refresh( (picFile,0) )
        
#----------------------------------------------------------------------------
# At the moment, one run of this app is hardcoded to stay in one category
# of items, eg. Vinyl vs Bequties vs Stamps. Later, this should be a 
# GUI selectable filter


qtApp = QApplication( sys.argv )


qWin = uiLayout("/home/volker/Dropbox/private_python/ebay", qtApp ) 
qWin.show()

qtApp.exec()

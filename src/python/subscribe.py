# coding: utf-8
import os
import uno
import unohelper
from com.sun.star.task import XJob

import json
import ssl
from urllib import request
import urllib
import pathlib
import webbrowser
import traceback
import csv
import time
from com.sun.star.ui.dialogs.TemplateDescription import FILESAVE_SIMPLE
from subscribe_utils import msgbox, getProjectDataPath, createUnoService
import base64

class SubscribeImp(unohelper.Base, XJob):
    def __init__(self, ctx):
        self.ctx = ctx

    ## http request
    #  @param url string
    def makeReq(self, url):
        try:
            context = ssl._create_unverified_context()
            result = urllib.request.urlopen(url, timeout=2, context=context)
            # ~ msgbox(result.getcode())
        except urllib.error.HTTPError as e:
            msgbox("e.code")
        return result

    ## web browser (Default browser)
    #  @param webfilename string
    def useWebBrowser(self, webfilename):
        try:
            oURL = webfilename
            if webbrowser.open(oURL,new=1,autoraise=1)== False:		# false is error
                raise webbrowser.Error
            oDisp = 'Success'
        except webbrowser.Error:
            oDisp = str(sys.stderr) + ' open failed'
        except:
            oDisp = traceback.format_exc(sys.exc_info()[2])

    ## export csv
    #  @param
    def exportCSV(self):
        try:
            # ToDo: oTableName use conf file
            oTableName = 'SUBSCRIBEDATA'
            context = createUnoService('com.sun.star.sdb.DatabaseContext')

            oCtx = uno.getComponentContext()
            oServiceManager = oCtx.ServiceManager
            oFilePicker = oServiceManager.\
                createInstanceWithArgumentsAndContext('com.sun.star.ui.dialogs.FilePicker',(FILESAVE_SIMPLE,),oCtx)
            # 預設存檔目錄
            oFilePicker.setDisplayDirectory('c:/temp')
            # Dialog標題設定
            oFilePicker.Title = "Export as"
            # 設定檔案類型
            oFilePicker.appendFilter("TEXT CSV(*.csv)","*.csv")
            oAccept = oFilePicker.execute()

            if oAccept == 1:
                oFiles = oFilePicker.getFiles()
                oFileURL = oFiles[0]
                oToFile = unohelper.fileUrlToSystemPath(oFileURL)
                if not oFileURL.endswith('.csv'):
                    oToFile = oToFile + '.csv'
            else:
                oDisp = '請選擇檔案匯出~'
                title="檔案訂閱通知"
                return

            oRst_BaseFile = getProjectDataPath() + "SubscribeData.odb"

            db = context.getByName(unohelper.systemPathToFileUrl(oRst_BaseFile))
            oCon = db.getConnection('','')
            oStm = oCon.createStatement()
            # RowSet
            oRowSet = createUnoService('com.sun.star.sdb.RowSet')
            oRowSet.ActiveConnection = oCon

            oSQL = "SELECT * FROM " + oTableName
            oRowSet.Command = oSQL
            oRowSet.execute()
            with open(oToFile, 'w', encoding='UTF8', newline='') as f:
                writer = csv.writer(f)
                while oRowSet.next():
                    data = [base64.b64encode(oRowSet.getString(1).encode('UTF-8')).decode('UTF-8'),
                            base64.b64encode(oRowSet.getString(2).encode('UTF-8')).decode('UTF-8'),
                            base64.b64encode(oRowSet.getString(3).encode('UTF-8')).decode('UTF-8'),
                            base64.b64encode(oRowSet.getString(4).encode('UTF-8')).decode('UTF-8'),
                            base64.b64encode(oRowSet.getString(5).encode('UTF-8')).decode('UTF-8'),
                            base64.b64encode(oRowSet.getString(6)[:-10].encode('UTF-8')).decode('UTF-8')]
                            # ~ oRowSet.getString(1),
                            # ~ oRowSet.getString(2),
                            # ~ oRowSet.getString(3),
                            # ~ oRowSet.getString(4),
                            # ~ oRowSet.getString(5),
                            # ~ oRowSet.getString(6)[:-10]
                    # write the data
                    writer.writerow(data)

            # Close Rowset
            oRowSet.close()

            # Unconnect with the Datasource
            oCon.close()

            oDisp = '已匯出檔案'
            title = "檔案訂閱通知"
        except Exception as er:
            oDisp = ''
            oDisp = str(traceback.format_exc()) + '\n' + str(er)
            title = 'Error'
        finally:
            msgbox(oDisp,title)
        return

    ## import csv
    #  @param
    def importCSV(self):
        try:
            oTableName = 'SUBSCRIBEDATA'
            context = createUnoService('com.sun.star.sdb.DatabaseContext')

            oCtx = uno.getComponentContext()
            oServiceManager = oCtx.ServiceManager
            oFilePicker = oServiceManager.createInstance('com.sun.star.ui.dialogs.FilePicker')
            # 預設存檔目錄
            oFilePicker.setDisplayDirectory('c:/temp')
            # Dialog標題設定
            oFilePicker.Title = "檔案選取"
            # 設定檔案類型
            oFilePicker.appendFilter("TEXT CSV(*.csv)","*.csv")
            oAccept = oFilePicker.execute()
            oDisp = ''

            if oAccept == 1:
                oFiles = oFilePicker.getFiles()
                oFileURL = oFiles[0]
                oToFile = unohelper.fileUrlToSystemPath(oFileURL)
                # ~ oDisp = oDisp + '\n\n' + str(oToFile) + '\n\n '
            else:
                oDisp = '請選擇檔案匯入~'
                title="檔案訂閱通知"
                return 0
            
            # ~ oRst_BaseFile = 'C:\\Users\\alantom\\AppData\\Roaming\\Subscription\\Data\\SubscribeData.odb'
            oRst_BaseFile = getProjectDataPath() + "SubscribeData.odb"
            oExcelCsv = str(oToFile)

            db = context.getByName(unohelper.systemPathToFileUrl(oRst_BaseFile))
            oCon = db.getConnection('','')
            oStm = oCon.createStatement()
            # RowSet
            oRowSet = createUnoService('com.sun.star.sdb.RowSet')
            oRowSet.ActiveConnection = oCon

            with open(oExcelCsv, 'r', encoding="utf-8", newline='') as oOpenObj:
                csvObj = csv.reader(oOpenObj, delimiter=',')
                index = 0
                # INSERT DATA
                for data in csvObj:
                    index = index + 1
                    uuid = base64.b64decode(data[1]).decode('UTF-8')
                    oSQLChk1 = "SELECT * FROM " + oTableName + " WHERE UUID = '" + uuid + "'"
                    oRowSet.Command = oSQLChk1
                    oRowSet.execute()
                    # ~ # ToDo: 驗證匯入data是否正確
                    if oRowSet.RowCount == 0:
                        oVal = "VALUES('" + \
                                    base64.b64decode(data[1]).decode('UTF-8') + "','" + \
                                    base64.b64decode(data[2]).decode('UTF-8') + "','" + \
                                    base64.b64decode(data[3]).decode('UTF-8') + "','" + \
                                    base64.b64decode(data[4]).decode('UTF-8') + "','" + \
                                    base64.b64decode(data[5]).decode('UTF-8') + "');"
                                    # ~ str(data[1]) + "','" + \
                                    # ~ str(data[2]) + "','" + \
                                    # ~ str(data[3]) + "','" + \
                                    # ~ str(data[4]) + "','" + \
                                    # ~ str(data[5]) + "');"
                        # ~ msgbox(oVal)
                        oSQL2 = "INSERT INTO " + oTableName + " (UUID,FILENAME,URL,SERVERNAME,RECORDTIME) " + oVal
                        oStm.executeUpdate(oSQL2)

            # Base Document Save
            db.DatabaseDocument.store()
            
            # Close Rowset
            oRowSet.close()

            # Unconnect with the Datasource
            oCon.close()

            oDisp = '檔案匯入成功~'
            title = "檔案訂閱通知"
        except Exception as er:
            oDisp = ''
            oDisp = str(traceback.format_exc()) + '\n' + str(er)
            title = 'Error'
        finally:
            msgbox(oDisp,title)
        return 1

    def execute(self, args):
        for prop in args:
            if prop.Name == 'GetApi':
                # ~ msgbox("SubscribeImp GetApi")
                url = prop.Value
                url = base64.b64decode(url).decode('UTF-8')
                res = self.makeReq(url)
                jsondata = json.loads(res.read().decode())
                data = []
                data.append(jsondata['0']['uuid'])
                data.append(jsondata['0']['filename'])
                data.append(jsondata['0']['url'])
                data.append(jsondata['0']['timestamp'])
                data.append(jsondata['0']['servername'])
                # ~ data.append(base64.b64encode(jsondata['0']['url'].encode('UTF-8')))
                # ~ data.append(base64.b64encode(jsondata['0']['servername'].encode('UTF-8')))
                return data

            if prop.Name == 'GetAllApi':
                # ~ msgbox("SubscribeImp GetAllApi")
                url = prop.Value
                res = self.makeReq(url)
                jsondata = json.loads(res.read().decode())
                # ~ data = [[],[],[],[],[]]
                data=[]
                for j in range(len(jsondata)):
                    col = []
                    data.append(col)

                for i in jsondata:
                    data[int(i)].append(jsondata[i]['uuid'])
                    data[int(i)].append(jsondata[i]['filename'])
                    data[int(i)].append(jsondata[i]['url'])
                    data[int(i)].append(jsondata[i]['timestamp'])
                    data[int(i)].append(jsondata[i]['servername'])
                return data

            if prop.Name == 'Export':
                # ~ msgbox("SubscribeImp Export")
                self.exportCSV()
                return

            if prop.Name == 'Import':
                # ~ msgbox("SubscribeImp Import")
                ret = self.importCSV()
                return ret

            if prop.Name == 'UseWebBrowser':
                # ~ msgbox("SubscribeImp UseWebBrowser")
                webfilename = prop.Value
                self.useWebBrowser(webfilename)
                return

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(SubscribeImp,
                                         "tw.ossii.Subscription.impl",
                                         ("com.sun.star.task.Job",),)

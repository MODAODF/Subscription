# -*- coding: utf-8 -*-

import uno
import unohelper
import os
from com.sun.star.beans import PropertyValue

# ------------HELPERS----------------
def createUnoService(service, ctx=None, args=None):
    '''
    Instanciate a Uno service.

    @service: name of the service to be instanciated.
    @ctx: the context if required.
    @args: the arguments when needed.
    '''
    if not ctx:
        ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    if ctx and args:
        return smgr.createInstanceWithArgumentsAndContext(service, args, ctx)
    elif args:
        return smgr.createInstanceWithArguments(service, args)
    elif ctx:
        return smgr.createInstanceWithContext(service, ctx)
    else:
        return smgr.createInstance(service)


def getConfigurationAccess(nodevalue, updatable=False):
    '''
    Access configuration value.

    @nodevalue: the configuration key node as a string.
    @updatable: set True when accessor needs to modify the key value.
    '''
    cp = createUnoService("com.sun.star.configuration.ConfigurationProvider")
    node = PropertyValue("nodepath", 0, nodevalue, 0)
    if updatable:
        return cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", (node,))
    else:
        return cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))

def getProductName():
    '''
    Return the program name.
    '''
    key = "/org.openoffice.Setup/Product"
    reader = getConfigurationAccess(key)
    return reader.ooName

def getProjectDataPath():
    appDataDir = os.getenv('APPDATA')
    outPath = appDataDir + "\\Subscription\\Data\\"
    if not os.path.exists(outPath):
        os.makedirs(outPath)

    return outPath

# ------------MSGBOX----------------
from com.sun.star.awt.MessageBoxType import (MESSAGEBOX, INFOBOX, ERRORBOX, WARNINGBOX, QUERYBOX)

def msgbox(message, title="Message", boxtype='message', buttons=1, win=None):
    '''
    Simple message box.

    Like the oobasic build-in function msgbox,
    but simplified as only intended for quick debugging.
    Signature: msgbox(message, title='Message', boxtype='message', buttons=1, win=None).
    '''
    types = {'message': MESSAGEBOX, 'info': INFOBOX, 'error': ERRORBOX,
             'warning': WARNINGBOX, 'query': QUERYBOX}
    tk = createUnoService("com.sun.star.awt.Toolkit")
    if not win:
        desktop = createUnoService("com.sun.star.frame.Desktop")
        frame = desktop.ActiveFrame
        if frame.ActiveFrame:
            # top window is a subdocument
            frame = frame.ActiveFrame
        win = frame.ComponentWindow
    box = tk.createMessageBox(win, types[boxtype], buttons, title, message)
    return box.execute()

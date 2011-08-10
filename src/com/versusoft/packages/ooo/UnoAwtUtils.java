/**
 *  odt2daisy - OpenDocument to DAISY XML/Audio
 *
 *  (c) Copyright 2008 - 2011 by Vincent Spiewak, All Rights Reserved.
 *
 *  This program is free software; you can redistribute it and/or modify
 *  it under the terms of the GNU General Lesser Public License as published by
 *  the Free Software Foundation; either version 3 of the License, or
 *  (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU Lesser General Public License for more details.
 *
 *  You should have received a copy of the GNU Lesser General Public License along
 *  with this program; if not, write to the Free Software Foundation, Inc.,
 *  51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
 */
package com.versusoft.packages.ooo;

import com.sun.star.awt.MessageBoxButtons;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.XMessageBox;
import com.sun.star.awt.XMessageBoxFactory;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.XPropertySet;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XInitialization;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.ui.dialogs.XExecutableDialog;
import com.sun.star.ui.dialogs.XFilePicker;
import com.sun.star.ui.dialogs.XFilterManager;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 * Utility class for user interface.
 * @author Vincent Spiewak
 */
public class UnoAwtUtils {

    /**
     * Utility function for displaying a Save as dialog.
     *
     * @param filename Initial file name shown in Save as dialog.
     * @param filterName Filter name for the dialog.
     * @param filterPattern Filter pattern for the dialog.
     * @param m_xContext Component context to be passed to a component via ::com::sun::star::lang::XSingleComponentFactory.
     * @return The path from the dialog.
     */
    public static String showSaveAsDialog(String filename, String filterName, String filterPattern, XComponentContext m_xContext) {
    //public static String showSaveAsDialog(String filename, String filterName, String filterPattern, String initialDirectory, XComponentContext m_xContext) { //@todo change initial output dir
        String sStorePath = "";
        XComponent xComponent = null;
        XMultiComponentFactory m_xMCF = m_xContext.getServiceManager();

        try {
            // the filepicker is instantiated with the global Multicomponentfactory...
            Object oFilePicker = m_xMCF.createInstanceWithContext("com.sun.star.ui.dialogs.FilePicker", m_xContext);
            XFilePicker xFilePicker = (XFilePicker) UnoRuntime.queryInterface(XFilePicker.class, oFilePicker);

            // choose the template that defines the capabilities of the filepicker dialog
            XInitialization xInitialize = (XInitialization) UnoRuntime.queryInterface(XInitialization.class, xFilePicker);
            Short[] listAny = new Short[]{new Short(com.sun.star.ui.dialogs.TemplateDescription.FILESAVE_AUTOEXTENSION)};
            xInitialize.initialize(listAny);

            // add a control to the dialog to add the extension automatically to the filename...
            // CRASH ON OOo Beta 3 MACOSX 
            //XFilePickerControlAccess xFilePickerControlAccess = (XFilePickerControlAccess) UnoRuntime.queryInterface(XFilePickerControlAccess.class, xFilePicker);
            //xFilePickerControlAccess.setValue(com.sun.star.ui.dialogs.ExtendedFilePickerElementIds.CHECKBOX_AUTOEXTENSION, (short) 0, new Boolean(true));

            xComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, xFilePicker);

            // execute the dialog...
            XExecutableDialog xExecutable = (XExecutableDialog) UnoRuntime.queryInterface(XExecutableDialog.class, xFilePicker);

            // set the filters of the dialog. The filternames may be retrieved from
            // http://wiki.services.openoffice.org/wiki/Framework/Article/Filter
            XFilterManager xFilterManager = (XFilterManager) UnoRuntime.queryInterface(XFilterManager.class, xFilePicker);
            xFilterManager.appendFilter(filterName, filterPattern);

            // set the initial display directory.
            Object oPathSettings = m_xMCF.createInstanceWithContext("com.sun.star.util.PathSettings", m_xContext);
            XPropertySet xPropertySet = (XPropertySet) com.sun.star.uno.UnoRuntime.queryInterface(XPropertySet.class, oPathSettings);
            String sInitialDir = (String) xPropertySet.getPropertyValue("Work_writable");
            xFilePicker.setDisplayDirectory(sInitialDir); //@todo change output dir to current document location!

            //set the initial filename
                xFilePicker.setDefaultName(filename);
            
            short nResult = xExecutable.execute();

            // query the resulting path of the dialog...
            if (nResult == com.sun.star.ui.dialogs.ExecutableDialogResults.OK) {
                String[] sPathList = xFilePicker.getFiles();
                if (sPathList.length > 0) {
                    sStorePath = sPathList[0];
                }
            }

        } catch (com.sun.star.uno.Exception exception) {
            exception.printStackTrace();

        } finally {
            //make sure always to dispose the component and free the memory!
            if (xComponent != null) {
                xComponent.dispose();
            }
        }
        return sStorePath;
    }

    /**
     * Displays a dialog, unless one of the parameters is null.
     *
     * @param parentWindowPeer The actual window implementation on the device.
     * @param messageBoxType The type of dialog (infobox, warningbox, errorbox, querybox or messbox).
     * @param messageBoxButtons A number that specifies which buttons should be available on the message box.
     * @param messageBoxTitle The title of the dialog.
     * @param message The text to be displayed in the dialog.
     * @return The execution of the dialog.
     */
    public static short showMessageBox(XWindowPeer parentWindowPeer, String messageBoxType, int messageBoxButtons, String messageBoxTitle, String message) {
        if (parentWindowPeer == null || messageBoxType == null || messageBoxTitle == null || message == null) {
            return 0;
        }

        // Initialize the message box factory
        XMessageBoxFactory messageBoxFactory = (XMessageBoxFactory) UnoRuntime.queryInterface(XMessageBoxFactory.class, parentWindowPeer.getToolkit());

        Rectangle messageBoxRectangle = new Rectangle();

        XMessageBox box = messageBoxFactory.createMessageBox(parentWindowPeer, messageBoxRectangle, messageBoxType, messageBoxButtons, messageBoxTitle, message);
        return box.execute();
    }

    /**
     * Displays an information dialog, unless one of the parameters is null.
     *
     * @param parentWindowPeer The actual window implementation on the device.
     * @param messageBoxTitle The title of the dialog.
     * @param message The text to be displayed in the dialog.
     * @return The information dialog.
     */
    public static short showInfoMessageBox(XWindowPeer parentWindowPeer, String messageBoxTitle, String message) {
        if (parentWindowPeer == null || messageBoxTitle == null || message == null) {
            return 0;
        }

        return showMessageBox(parentWindowPeer, "infobox", MessageBoxButtons.BUTTONS_OK, messageBoxTitle, message);
    }

    /**
     * Displays an error dialog, unless one of the parameters is null.
     * 
     * @param parentWindowPeer The actual window implementation on the device.
     * @param messageBoxTitle The title of the dialog.
     * @param message The text to be displayed in the dialog.
     * @return The error dialog.
     */
    public static short showErrorMessageBox(XWindowPeer parentWindowPeer, String messageBoxTitle, String message) {
        if (parentWindowPeer == null || messageBoxTitle == null || message == null) {
            return 0;
        }

        return showMessageBox(parentWindowPeer, "errorbox", MessageBoxButtons.BUTTONS_OK, messageBoxTitle, message);
    }

    /**
     * Displays a Yes/No confirmation dialog, unless one of the parameters is null.
     *
     * @param parentWindowPeer The actual window implementation on the device.
     * @param messageBoxTitle The title of the dialog.
     * @param message The text to be displayed in the dialog.
     * @return The Yes/No confirmation dialog.
     */
    public static short showYesNoWarningMessageBox(XWindowPeer parentWindowPeer, String messageBoxTitle, String message) {
        if (parentWindowPeer == null || messageBoxTitle == null || message == null) {
            return 0;
        }

        return showMessageBox(parentWindowPeer, "warningbox", MessageBoxButtons.BUTTONS_YES_NO + MessageBoxButtons.DEFAULT_BUTTON_NO, messageBoxTitle, message);
    }

    /**
     * Displays an OK/Cancel confirmation dialog, unless one of the parameters is null.
     *
     * @param parentWindowPeer The actual window implementation on the device.
     * @param messageBoxTitle The title of the dialog.
     * @param message The text to be displayed in the dialog.
     * @return The OK/Cancel dialog.
     */
    public static short showOkCancelWarningMessageBox(XWindowPeer parentWindowPeer, String messageBoxTitle, String message) {
        if (parentWindowPeer == null || messageBoxTitle == null || message == null) {
            return 0;
        }

        return showMessageBox(parentWindowPeer, "warningbox", MessageBoxButtons.BUTTONS_OK_CANCEL + MessageBoxButtons.DEFAULT_BUTTON_OK, messageBoxTitle, message);
    }

    /**
     * Displays a Yes/No/Cancel confirmation dialog, unless one of the parameters is null.
     *
     * @param parentWindowPeer The actual window implementation on the device.
     * @param messageBoxTitle The title of the dialog.
     * @param message The text to be displayed in the dialog.
     * @return The Yes/No/Cancel dialog.
     */
    public static short showQuestionMessageBox(XWindowPeer parentWindowPeer, String messageBoxTitle, String message) {
        if (parentWindowPeer == null || messageBoxTitle == null || message == null) {
            return 0;
        }

        return showMessageBox(parentWindowPeer, "querybox", MessageBoxButtons.BUTTONS_YES_NO_CANCEL + MessageBoxButtons.DEFAULT_BUTTON_YES, messageBoxTitle, message);
    }

    /**
     * Displays an Abort/Retry/Ignore dialog, unless one of the parameters is null.
     *
     * @param parentWindowPeer The actual window implementation on the device.
     * @param messageBoxTitle The title of the dialog.
     * @param message The text to be displayed in the dialog.
     * @return The Abort/Retry/Ignore dialog.
     */
    public static short showAbortRetryIgnoreErrorMessageBox(XWindowPeer parentWindowPeer, String messageBoxTitle, String message) {
        if (parentWindowPeer == null || messageBoxTitle == null || message == null) {
            return 0;
        }

        return showMessageBox(parentWindowPeer, "errorbox", MessageBoxButtons.BUTTONS_ABORT_IGNORE_RETRY + MessageBoxButtons.DEFAULT_BUTTON_RETRY, messageBoxTitle, message);
    }

    /**
     * Displays a Retry/Cancel dialog, unless one of the parameters is null.
     *
     * @param parentWindowPeer The actual window implementation on the device.
     * @param messageBoxTitle The title of the dialog.
     * @param message The text to be displayed in the dialog.
     * @return The Retry/Cancel dialog.
     */
    public static short showRetryCancelErrorMessageBox(XWindowPeer parentWindowPeer, String messageBoxTitle, String message) {
        if (parentWindowPeer == null || messageBoxTitle == null || message == null) {
            return 0;
        }

        return showMessageBox(parentWindowPeer, "errorbox", MessageBoxButtons.BUTTONS_RETRY_CANCEL + MessageBoxButtons.DEFAULT_BUTTON_CANCEL, messageBoxTitle, message);
    }
}

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
package com.versusoft.packages.ooo.odt2daisy.addon.gui;

import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.frame.XStorable;
import com.sun.star.task.XStatusIndicator;
import com.sun.star.task.XStatusIndicatorFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.versusoft.packages.ooo.UnoAwtUtils;
import com.versusoft.packages.ooo.UnoUtils;
import com.versusoft.packages.ooo.odt2daisy.Odt2Daisy;
import com.versusoft.packages.ooo.odt2daisy.addon.pipelinelite.PipelineLite;
import java.io.File;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.logging.FileHandler;
import java.util.logging.Handler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;
import org.apache.commons.io.FileSystemUtils;

public class UnoGUI {

    private static final String LOG_FILENAME = "odt2daisyaddon.log";
    private static final String TMP_ODT_PREFIX = "odt2daisy";
    private static final String TMP_ODT_SUFFIX = ".odt";
    private static final String FLAT_XML_FILTER_NAME = "writer8";
    private static final String IMAGE_DIR = "images/";
    private static final Logger logger = Logger.getLogger("com.versusoft.packages.ooo.odt2daisy.addon.gui");
    // L10N Strings
    private static String L10N_MessageBox_Warning_Title = null;
    private static String L10N_No_Headings_Warning = null;
    private static String L10N_Incompatible_Images_Error = null;
    private static String L10N_Default_Export_Filename = null;
    private static String L10N_MessageBox_Error_Title = null;
    private static String L10N_DTD_Error_Message = null;
    private static String L10N_Line = null;
    private static String L10N_Message = null;
    private static String L10N_MessageBox_InternalError_Title = null;
    private static String L10N_Export_Aborted_Message = null;
    private static String L10N_MessageBox_Info_Title = null;
    private static String L10N_Empty_Document_Message = null;
    private static String L10N_Validated_DTD_Message = null;
    private static String L10N_PipelineLite_Update_Message = null;
    private static String L10N_PipelineLite_Size_Error_Message = null;
    private static String L10N_PipelineLite_Extract_Message = null;
    private static String L10N_PipelineLite_Extract_Error_Message = null;
    private static String L10N_PipelineLite_Exec_Error_Message = null;
    private static String L10N_StatusIndicator_Step_1 = null;
    private static String L10N_StatusIndicator_Step_2 = null;
    private static String L10N_StatusIndicator_Step_3 = null;
    private static String L10N_StatusIndicator_Step_4 = null;
    private static String L10N_StatusIndicator_Step_5 = null;
    private static String L10N_StatusIndicator_Step_6 = null;
    private static String L10N_StatusIndicator_Step_7 = null;
    private static String L10N_StatusIndicator_Step_8 = null;
    private static String L10N_StatusIndicator_Step_9 = null;
    private static String L10N_StatusIndicator_Step_10 = null;

    private String exportUrl = null;
    private Locale OOoLocale = null;

    private XComponentContext m_xContext = null;
    private XFrame m_xFrame = null;
    private XModel xDoc = null;
    private XWindow parentWindow = null;
    private XWindowPeer parentWindowPeer = null;
    private XStatusIndicatorFactory xStatusIndicatorFactory = null;
    private XStatusIndicator xStatusIndicator = null;

    // Export dialog
    private ExportDialog dialog;

    // Logger fields
    private Handler fh = null;
    private File logFile = null;

    private boolean isFullExport = false;

    /**
     *
     * @param m_xContext Component context to be passed to a component via ::com::sun::star::lang::XSingleComponentFactory.
     * @param m_xFrame Frame object that serves as an "anchor" object where a component can be attached to.
     */
    public UnoGUI(XComponentContext m_xContext, XFrame m_xFrame){

        this(m_xContext, m_xFrame, false);

    }

    /**
     *
     * @param m_xContext Component context to be passed to a component via ::com::sun::star::lang::XSingleComponentFactory.
     * @param m_xFrame Frame object that serves as an "anchor" object where a component can be attached to.
     * @param isFullExport true if the content should be exported as Full DAISY, false if the content should be exported as DAISY XML (no audio).
     */
    public UnoGUI(XComponentContext m_xContext, XFrame m_xFrame, boolean isFullExport) {

        this.m_xContext = m_xContext;
        this.m_xFrame = m_xFrame;
        this.isFullExport = isFullExport;
        
        try {

            // Configuring logger
            logFile = File.createTempFile(LOG_FILENAME, null);
            fh = new FileHandler(logFile.getAbsolutePath());
            fh.setFormatter(new SimpleFormatter());
            Logger.getLogger("").addHandler(fh);
            Logger.getLogger("").setLevel(Level.FINEST);
            logger.fine("entering");

            // Configuring Locale
            OOoLocale = new Locale(UnoUtils.getUILocale(m_xContext));
            L10N_MessageBox_Warning_Title = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("MessageBox_Warning_Title");
            L10N_No_Headings_Warning = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("No_Headings_Warning");
            L10N_Incompatible_Images_Error = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("Incompatible_Images_Error");
            L10N_Default_Export_Filename = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("Default_Export_Filename");
            L10N_MessageBox_Error_Title = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("MessageBox_Error_Title");
            L10N_DTD_Error_Message = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("DTD_Error_Message");
            L10N_Line = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("Line");
            L10N_Message = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("Message");
            L10N_MessageBox_InternalError_Title = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("MessageBox_InternalError_Title");
            L10N_Export_Aborted_Message = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("Export_Aborted_Message");
            L10N_MessageBox_Info_Title = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("MessageBox_Info_Title");
            L10N_Empty_Document_Message = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("Empty_Document_Message");
            L10N_Validated_DTD_Message = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("Validated_DTD_Message");
            L10N_PipelineLite_Update_Message = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("PipelineLite_Update_Message");
            L10N_PipelineLite_Size_Error_Message = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("PipelineLite_Size_Error_Message");
            L10N_PipelineLite_Extract_Message = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("PipelineLite_Extract_Message");
            L10N_PipelineLite_Extract_Error_Message = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("PipelineLite_Extract_Error_Message");
            L10N_PipelineLite_Exec_Error_Message = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("PipelineLite_Exec_Error_Message");
            L10N_StatusIndicator_Step_1 = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("StatusIndicator_Step_1");
            L10N_StatusIndicator_Step_2 = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("StatusIndicator_Step_2");
            L10N_StatusIndicator_Step_3 = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("StatusIndicator_Step_3");
            L10N_StatusIndicator_Step_4 = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("StatusIndicator_Step_4");
            L10N_StatusIndicator_Step_5 = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("StatusIndicator_Step_5");
            L10N_StatusIndicator_Step_6 = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("StatusIndicator_Step_6");
            L10N_StatusIndicator_Step_7 = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("StatusIndicator_Step_7");
            L10N_StatusIndicator_Step_8 = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("StatusIndicator_Step_8");
            L10N_StatusIndicator_Step_9 = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("StatusIndicator_Step_9");
            L10N_StatusIndicator_Step_10 = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("StatusIndicator_Step_10");

            // Init Status Indicator
            xStatusIndicatorFactory = (XStatusIndicatorFactory) UnoRuntime.queryInterface(XStatusIndicatorFactory.class, m_xFrame);
            xStatusIndicator = xStatusIndicatorFactory.createStatusIndicator();

            // Query Uno Object
            xDoc = (XModel) UnoRuntime.queryInterface(
                    XModel.class, m_xFrame.getController().getModel());

            parentWindow = xDoc.getCurrentController().getFrame().getContainerWindow();
            parentWindowPeer = (XWindowPeer) UnoRuntime.queryInterface(XWindowPeer.class, parentWindow);

            //DEBUG MODE
            //UnoAwtUtils.showInfoMessageBox(parentWindowPeer, L10N_MessageBox_Info_Title, "DEBUG MODE: "+logFile.getAbsolutePath());
            
        } catch (Exception ex) {
            logger.log(Level.SEVERE, null, ex);
        }

    }

    /**
     * Save as DAISY XML.
     *
     * @return true if the the content could be saved as valid DAISY XML, false otherwise (e.g. an error occurred).
     */
    public boolean saveAsXML() {

        Odt2Daisy odt2daisy = null;

        String tmpOdtUrl = null;
        String tmpOdtUnoUrl = null;
        String exportUnoUrl = null;

        try {

            // Start status bar
            xStatusIndicator.start(L10N_StatusIndicator_Step_1, 100);
            xStatusIndicator.setValue(5);

            // Request a temporary file
            xStatusIndicator.setText(L10N_StatusIndicator_Step_2);
            xStatusIndicator.setValue(10);

            logger.fine("request a temporary file");

            File tmpFile = File.createTempFile(
                    TMP_ODT_PREFIX,
                    TMP_ODT_SUFFIX);
            tmpFile.deleteOnExit();
            tmpOdtUrl = tmpFile.getAbsolutePath();
            tmpOdtUnoUrl =
                    UnoUtils.createUnoFileURL(tmpOdtUrl, m_xContext);

            logger.fine("tmpOdtUrl:" + tmpOdtUrl);
            logger.fine("tmpOdtUnoUrl:" + tmpOdtUnoUrl);


            // Export in ODT Format using UNO API
            xStatusIndicator.setText(L10N_StatusIndicator_Step_3);
            xStatusIndicator.setValue(15);

            logger.fine("save current document in ODT using UNO API");

            PropertyValue[] conversionProperties = new PropertyValue[1];
            conversionProperties[0] = new PropertyValue();
            conversionProperties[0].Name = "FilterName";
            conversionProperties[0].Value = FLAT_XML_FILTER_NAME; //Daisy DTBook OpenDocument XML

            XStorable storable = (XStorable) UnoRuntime.queryInterface(
                    XStorable.class, m_xFrame.getController().getModel());

            storable.storeToURL(tmpOdtUnoUrl, conversionProperties);


            // Create and Init odt2daisy
            xStatusIndicator.setText(L10N_StatusIndicator_Step_4);
            xStatusIndicator.setValue(20);

            logger.fine("create and init odt2daisy");
            odt2daisy = new Odt2Daisy(tmpOdtUrl); //@todo add initial output directory URL
            odt2daisy.init();

            xStatusIndicator.setText(L10N_StatusIndicator_Step_5);
            xStatusIndicator.setValue(40);

            // Stop Progress bar during user inputs
            xStatusIndicator.end();

            // Show an alert if ODT empty
            if (odt2daisy.isEmptyDocument()) {

                String messageBoxTitle =
                        L10N_MessageBox_Info_Title;
                String message =
                        "\n" +
                        L10N_Empty_Document_Message + "    \n";

                UnoAwtUtils.showInfoMessageBox(parentWindowPeer, messageBoxTitle, message);
                return false;

            }

            // Show a warning if ODT does not contain any headings (Heading 1)
            if (!odt2daisy.isUsingHeadings()) {

                String messageBoxTitle =
                        L10N_MessageBox_Warning_Title;
                String message =
                        L10N_No_Headings_Warning + "\n";

                Short result = UnoAwtUtils.showYesNoWarningMessageBox(parentWindowPeer, messageBoxTitle, message);

                // Abort on Cancel
                if (result == (short) 3) {
                    logger.fine("User cancelled export");
                    return false;
                }
            }

            // @todo Ideally, warning if ODT contains images in format not supported by DAISY 3 should be here, instead of after the Save as dialog.


            // Raise File Export Dialog @todo add initial output directory URL
            exportUnoUrl = UnoAwtUtils.showSaveAsDialog(L10N_Default_Export_Filename, "DAISY DTBook XML", "*.xml", m_xContext);
            logger.fine("exportUnoUrl=" + exportUnoUrl);

            if (exportUnoUrl.length() < 1) {
                logger.info("user cancelled export");
                return false;
            }


            // Auto append extension manually because crash autoextension crash on OOo beta 3 macosx
            if (!exportUnoUrl.endsWith(".xml")) {
                exportUnoUrl = exportUnoUrl.concat(".xml");
            }


            exportUrl = UnoUtils.UnoURLtoURL(exportUnoUrl, m_xContext);
            logger.fine("exportUrl=" + exportUrl);

            // Raise Export Dialog Options
            dialog = new ExportDialog(m_xContext, isFullExport);
            dialog.setUid(odt2daisy.getUidParam());
            dialog.setDoctitle(odt2daisy.getTitleParam());
            dialog.setCreator(odt2daisy.getCreatorParam());
            dialog.setPublisher(odt2daisy.getPublisherParam());
            dialog.setProducer(odt2daisy.getProducerParam());
            dialog.setLang(odt2daisy.getLangParam());
            dialog.setAlternateLevelMarkup(odt2daisy.isUseAlternateLevelParam());

            boolean retDialog = dialog.execute();
            if (!retDialog) {
                logger.info("user cancelled export");
                return false;
            }

            xStatusIndicator.start(L10N_StatusIndicator_Step_6, 100);
            xStatusIndicator.setValue(45);

            if (dialog.isPaginationEnable()) {
                logger.info("Pagination process started");
                odt2daisy.paginationProcessing();
                logger.info("Pagination process end");

            }

            // Correction Processing
            xStatusIndicator.setText(L10N_StatusIndicator_Step_7);
            xStatusIndicator.setValue(60);
            
            logger.fine("Trying ODF XML correction");// was: "daisy correction"??
            odt2daisy.correctionProcessing();

            // Set Params according to DAISY Expport dialog
            odt2daisy.setUidParam(dialog.getUid());
            odt2daisy.setTitleParam(dialog.getDoctitle());
            odt2daisy.setCreatorParam(dialog.getCreator());
            odt2daisy.setPublisherParam(dialog.getPublisher());
            odt2daisy.setProducerParam(dialog.getProducer());
            odt2daisy.setUseAlternateLevelParam(dialog.isAlternateLevelMarkup());
            odt2daisy.setWriteCSSParam(dialog.isWriteCSS());

            // Convert as DAISY XML
            xStatusIndicator.setText(L10N_StatusIndicator_Step_8);
            xStatusIndicator.setValue(70);

            logger.fine("Trying daisy translation");
            odt2daisy.convertAsDTBook(exportUrl, IMAGE_DIR);

            //@todo Check whether imagesProcessing(dtbookFile, imageDir) [in odt2daisy object] can be moved to a method that is called before odt2daisy.convertAsDTBook(exportUrl, IMAGE_DIR)
            // or just check merged XML instead (i.e. without checking file names stored inside ODF).
            if (odt2daisy.containsIncompatibleImages()) {
                LinkedHashMap<String, String> incompatibleImg = odt2daisy.getIncompatibleImages();
                StringBuffer messageWithImgList = new StringBuffer(L10N_Incompatible_Images_Error);
                for (java.util.Map.Entry<String,String> entry: incompatibleImg.entrySet()) {
                    messageWithImgList.append("* ").append(entry.getKey()).append("\n");
                }
                String message = messageWithImgList.toString();
                Short result = UnoAwtUtils.showErrorMessageBox(parentWindowPeer, L10N_MessageBox_Error_Title, message);
                logger.severe(message);
                return false;
            }

            // DTD Validation
            xStatusIndicator.setText(L10N_StatusIndicator_Step_9);
            xStatusIndicator.setValue(90);

            logger.fine("Trying daisy DTD validation");
            odt2daisy.validateDTD(exportUrl);

            if (odt2daisy.getErrorHandler().hadError()) {

                String messageBoxTitle =
                        L10N_MessageBox_Error_Title;
                String message =
                        L10N_DTD_Error_Message + "\n\n" +
                        L10N_Line + ": " + odt2daisy.getErrorHandler().getLineNumber() + "\n" +
                        L10N_Message + ": " + odt2daisy.getErrorHandler().getMessage() + "\n" +
                        "\n";

                UnoAwtUtils.showErrorMessageBox(parentWindowPeer, messageBoxTitle, message);
                logger.severe(message);
                return false;
            }

            xStatusIndicator.setText(L10N_StatusIndicator_Step_10);
            xStatusIndicator.setValue(100);

            return true;

        } catch (Exception e) {

            String messageBoxTitle = L10N_MessageBox_InternalError_Title;
            String message =
                    L10N_Export_Aborted_Message + " " +
                    logFile.getAbsolutePath() + "\n";

            UnoAwtUtils.showErrorMessageBox(parentWindowPeer, messageBoxTitle, message);

            if (logger != null) {
                logger.log(Level.SEVERE, null, e);
            }

            return false;

        }

    }

    /**
     * Save as Full DAISY.
     */
    public void saveAsFull() {

        try {

            boolean isInUserFolder = false;
            boolean isInTmpFolder = false;
            boolean isExtracted = false;
            boolean isOutdated = false;

            String extract_dir = null;

            PipelineLite pipelinelite = new PipelineLite(logger);

            // check PipelineLite presence
            isInUserFolder = PipelineLite.isExctracted(PipelineLite.USERMODE_FOLDER);
            isInTmpFolder = PipelineLite.isExctracted(PipelineLite.TMPMODE_FOLDER);
            isExtracted = isInUserFolder || isInTmpFolder;


            // check if Pipeline is not outdated
            // set extract_dir
            if (isExtracted) {
                if (isInUserFolder) {
                    isOutdated = PipelineLite.isOutdated(PipelineLite.USERMODE_FOLDER);
                    extract_dir = PipelineLite.USERMODE_FOLDER;
                } else if (isInTmpFolder) {
                    isOutdated = PipelineLite.isOutdated(PipelineLite.TMPMODE_FOLDER);
                    extract_dir = PipelineLite.TMPMODE_FOLDER;

                }
            }

            logger.info("DEBUG extracted:"+isExtracted);
            logger.info("DEBUG outdated:"+isOutdated);
            logger.info("DEBUG in home:"+isInUserFolder);
            logger.info("DEBUG in tmpdir:"+isInTmpFolder);

            // show outdated warning
            if (isOutdated) {

                UnoAwtUtils.showInfoMessageBox(
                        parentWindowPeer,
                        L10N_MessageBox_Info_Title,
                        L10N_PipelineLite_Update_Message
                        );

                // force install update
                isExtracted = false;
            }

            // Extract PipelineLite
            if (!isExtracted) {


                // choose user home folder if enough disk space, else tmp
                if (FileSystemUtils.freeSpaceKb(PipelineLite.USERMODE_FOLDER) > PipelineLite.REQUIRED_KB) {
                    extract_dir = PipelineLite.USERMODE_FOLDER;
                } else if (FileSystemUtils.freeSpaceKb(PipelineLite.TMPMODE_FOLDER) > PipelineLite.REQUIRED_KB) {
                    extract_dir = PipelineLite.TMPMODE_FOLDER;
                } else {

                    UnoAwtUtils.showErrorMessageBox(
                            parentWindowPeer,
                            L10N_MessageBox_Error_Title,
                            L10N_PipelineLite_Size_Error_Message + " " +
                            (PipelineLite.REQUIRED_KB / 1000) + "Mo");
                    return;
                }

                // show dialog with extract folder info
                // avoid dialog on update
                if(!isOutdated){
                    UnoAwtUtils.showInfoMessageBox(
                            parentWindowPeer,
                            L10N_MessageBox_Info_Title,
                            L10N_PipelineLite_Extract_Message + " " +
                            extract_dir + PipelineLite.PIPELINE_FOLDER_NAME + " ");
                }

                // extract pipeline
                try {


                    logger.info("DEBUG: start extraction"+extract_dir);
                    pipelinelite.extract(extract_dir);
                    logger.info("DEBUG: end extraction"+extract_dir);

                } catch (Exception e) {

                    String message = L10N_PipelineLite_Extract_Error_Message + "\n";
                    message += L10N_Export_Aborted_Message
                            + " "
                            + logFile.getAbsolutePath();

                    e.printStackTrace();
                    logger.log(Level.SEVERE, message);

                    UnoAwtUtils.showErrorMessageBox(
                            parentWindowPeer,
                            L10N_MessageBox_Error_Title,
                            message);
                    return;
                }
            }

            // launch narrator
            try {

                int retcode = pipelinelite.launchNarrator(
                        extract_dir,
                        exportUrl,
                        exportUrl + ".full",
                        dialog.isFixRoutine(),
                        dialog.isSentDetection(),
                        dialog.getBitrate());

                logger.info("DEBUG retcode:"+retcode);

                if(retcode == 0){

                    //showOkSaveAsXML();
                    logger.info("Save as Full DAISY OK !");

                } else {
                   // throw new Exception();
                }

            } catch (Exception e) {

                String message = L10N_PipelineLite_Exec_Error_Message + "\n";
                message += L10N_Export_Aborted_Message + " "+ logFile.getAbsolutePath();

                logger.log(Level.SEVERE, message, e);
                e.printStackTrace();

                UnoAwtUtils.showErrorMessageBox(
                        parentWindowPeer,
                        L10N_MessageBox_Error_Title,
                        message);
                return;
            }

        } catch (Exception e) {

            String message =
                    L10N_Export_Aborted_Message +
                    " " + logFile.getAbsolutePath() + "\n";

            e.printStackTrace();
            logger.log(Level.SEVERE, message,e);

            UnoAwtUtils.showErrorMessageBox(parentWindowPeer,
                    L10N_MessageBox_InternalError_Title,
                    message);

            return;
        }

    }

    /**
     * Display information dialog saying that the content was exported as valid DAISY.
     */
    public void showOkSaveAsXML() {

        // Valid DTD Message
        String messageBoxTitle =
                L10N_MessageBox_Info_Title;
        String message =
                "\n" +
                L10N_Validated_DTD_Message + "\n";

        UnoAwtUtils.showInfoMessageBox(parentWindowPeer, messageBoxTitle, message);

    }

    /**
     * Flush any buffered output for the logger.
     *
     * @param fh The Handler object takes log messages from the Logger and exports them.
     */
    private static void flushLogger(Handler fh) {
        if (fh != null) {
            fh.flush();
            fh.close();
        }
    }

    /**
     * Flush any buffered output for the logger.
     */
    public void flushLogger(){
        flushLogger(fh);
    }

    /**
     * Stops the progress indicator.
     * @see http://api.openoffice.org/docs/common/ref/com/sun/star/task/XStatusIndicator.html
     */
    public void stopStatusIndicator() {
        if(xStatusIndicator != null){
            xStatusIndicator.end();
        }
    }

}

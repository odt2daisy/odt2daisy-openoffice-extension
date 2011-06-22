/**
 *  odt2daisy - OpenDocument to DAISY XML/Audio
 *
 *  (c) Copyright 2008 - 2009 by Vincent Spiewak, All Rights Reserved.
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

import com.sun.star.awt.PushButtonType;
import com.sun.star.awt.XCheckBox;
import com.sun.star.awt.XComboBox;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XControlModel;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameContainer;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.versusoft.packages.ooo.UnoUtils;
import java.util.Locale;

public class ExportDialog {

    // UNO Components
    private Object dialog;
    private Object dialogModel;
    private XDialog xDialog;
    private XComponent xComponent;
    private XControlContainer xControlCont;
    private XNameContainer xNameCont;
    // Buttons
    private static String _okButtonName = "Button1";
    private static String _cancelButtonName = "Button2";
    // Labels
    private static String _uidLabelName = "Label1";
    private static String _doctitleLabelName = "Label2";
    private static String _creatorLabelName = "Label3";
    private static String _publisherLabelName = "Label4";
    private static String _langLabelName = "Label5";
    private static String _producerLabelName = "Label6";
    private static String _bitrateLabelName = "Label7";
    // Fields
    private static String _uidFieldName = "Field1";
    private static String _doctitleFieldName = "Field2";
    private static String _creatorFieldName = "Field3";
    private static String _publisherFieldName = "Field4";
    private static String _langFieldName = "Field5";
    private static String _producerFieldName = "Field6";
    // CheckBox
    private static String _levelCBoxName = "CBox1";
    private static String _pageCBoxName = "CBox2";
    private static String _cssCBoxName = "CBox3";
    private static String _fixRoutineCBoxName = "CBox4";
    private static String _sentDetectionCBoxName = "CBox5";
    // bitrate control
    private static String _bitrateControlName = "bitrate";
    private static Object bitrateControlModel = null;
    // Localized Strings
    private static String L10N_okButtonValue = null;
    private static String L10N_cancelButtonValue = null;
    private static String L10N_uidLabelValue = null;
    private static String L10N_doctitleLabelValue = null;
    private static String L10N_creatorLabelValue = null;
    private static String L10N_publisherLabelValue = null;
    private static String L10N_langLabelValue = null;
    private static String L10N_levelLabelValue = null;
    private static String L10N_pageLabelValue = null;
    private static String L10N_cssLabelValue = null;
    private static String L10N_producerLabelValue = null;
    private static String L10N_titleDialogValue = null;
    private static String L10N_bitrateLabelValue = null;
    private static String L10N_fixRoutineLabelValue = null;
    private static String L10N_sentDetectionLabelValue = null;
    // controls values fields (for accesors)
    private String uid = null;
    private String doctitle = null;
    private String creator = null;
    private String publisher = null;
    private String producer = null;
    private String lang = null;
    private int bitrate = 0;
    private boolean alternateLevelMarkup = false;
    private boolean isPageEnabled = true;
    private boolean writeCSS = false;
    private boolean sentDetection = true;
    private boolean fixRoutine = true;
    // in save as full daisy
    private boolean isFullExport = false;

    public ExportDialog(XComponentContext m_xContext) {
        this(m_xContext, false);
    }

    public ExportDialog(XComponentContext m_xContext, boolean isFullExport) {

        this.isFullExport = isFullExport;

        try {

            /* Init Localized String */
            Locale OOoLocale = new Locale(UnoUtils.getUILocale(m_xContext));
            L10N_okButtonValue = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("Translate");
            L10N_cancelButtonValue = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("Cancel");
            L10N_uidLabelValue = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("UID") + " : ";
            L10N_doctitleLabelValue = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("Title") + " : ";
            L10N_creatorLabelValue = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("Creator") + " : ";
            L10N_publisherLabelValue = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("Publisher") + " : ";
            L10N_producerLabelValue = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("Producer") + " : ";
            L10N_langLabelValue = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("Language") + " : ";
            L10N_levelLabelValue = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("use_alternate_level_markup");
            L10N_pageLabelValue = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("include_page_numbers");
            L10N_cssLabelValue = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("include_CSS");
            L10N_titleDialogValue = new String(java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("DAISY_DTBook_Translator"));
            L10N_bitrateLabelValue = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("Bitrate") + " : ";
            L10N_fixRoutineLabelValue = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("fixRoutine");
            L10N_sentDetectionLabelValue = java.util.ResourceBundle.getBundle("com/versusoft/packages/ooo/odt2daisy/addon/l10n/Bundle", OOoLocale).getString("sentDetection");

            XMultiComponentFactory m_xMCF = m_xContext.getServiceManager();


            // Init Uno Dialog Control
            dialogModel = m_xMCF.createInstanceWithContext(
                    "com.sun.star.awt.UnoControlDialogModel", m_xContext);

            // dialog height
            Integer dialogHeight;
            if (isFullExport) {
                dialogHeight = new Integer(210);
            } else {
                dialogHeight = new Integer(175);
            }

            XPropertySet xPSetDialog = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, dialogModel);
            xPSetDialog.setPropertyValue("PositionX", new Integer(100));
            xPSetDialog.setPropertyValue("PositionY", new Integer(50));
            xPSetDialog.setPropertyValue("Width", new Integer(175));
            xPSetDialog.setPropertyValue("Height", dialogHeight);
            xPSetDialog.setPropertyValue("Title", L10N_titleDialogValue);


            // get the service manager from the dialog model
            XMultiServiceFactory xMultiServiceFactory = (XMultiServiceFactory) UnoRuntime.queryInterface(
                    XMultiServiceFactory.class, dialogModel);


            // Init UID Label
            Object uidLabelModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlFixedTextModel");

            XPropertySet xPSetLabel1 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, uidLabelModel);

            xPSetLabel1.setPropertyValue("PositionX", new Integer(14));
            xPSetLabel1.setPropertyValue("PositionY", new Integer(17));//63

            xPSetLabel1.setPropertyValue("Width", new Integer(35));
            xPSetLabel1.setPropertyValue("Height", new Integer(9));
            xPSetLabel1.setPropertyValue("Name", _uidLabelName);
            xPSetLabel1.setPropertyValue("Label", L10N_uidLabelValue);


            // Init Doctitle Label
            Object doctitleLabelModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlFixedTextModel");

            XPropertySet xPSetLabel2 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, doctitleLabelModel);

            xPSetLabel2.setPropertyValue("PositionX", new Integer(14));
            xPSetLabel2.setPropertyValue("PositionY", new Integer(33));
            xPSetLabel2.setPropertyValue("Width", new Integer(35));
            xPSetLabel2.setPropertyValue("Height", new Integer(9));
            xPSetLabel2.setPropertyValue("Name", _doctitleLabelName);
            xPSetLabel2.setPropertyValue("Label", L10N_doctitleLabelValue);


            // Init Creator Label
            Object creatorLabelModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlFixedTextModel");

            XPropertySet xPSetLabel3 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, creatorLabelModel);

            xPSetLabel3.setPropertyValue("PositionX", new Integer(14));
            xPSetLabel3.setPropertyValue("PositionY", new Integer(48));
            xPSetLabel3.setPropertyValue("Width", new Integer(35));
            xPSetLabel3.setPropertyValue("Height", new Integer(9));
            xPSetLabel3.setPropertyValue("Name", _creatorLabelName);
            xPSetLabel3.setPropertyValue("Label", L10N_creatorLabelValue);


            // Init Publisher Label
            Object publisherLabelModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlFixedTextModel");

            XPropertySet xPSetLabel4 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, publisherLabelModel);

            xPSetLabel4.setPropertyValue("PositionX", new Integer(14));
            xPSetLabel4.setPropertyValue("PositionY", new Integer(63));
            xPSetLabel4.setPropertyValue("Width", new Integer(35));
            xPSetLabel4.setPropertyValue("Height", new Integer(9));
            xPSetLabel4.setPropertyValue("Name", _publisherLabelName);
            xPSetLabel4.setPropertyValue("Label", L10N_publisherLabelValue);


            // Init Producer Label
            Object producerLabelModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlFixedTextModel");

            XPropertySet xPSetLabel5 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, producerLabelModel);

            xPSetLabel5.setPropertyValue("PositionX", new Integer(14));
            xPSetLabel5.setPropertyValue("PositionY", new Integer(78));
            xPSetLabel5.setPropertyValue("Width", new Integer(35));
            xPSetLabel5.setPropertyValue("Height", new Integer(9));
            xPSetLabel5.setPropertyValue("Name", _producerLabelName);
            xPSetLabel5.setPropertyValue("Label", L10N_producerLabelValue);


            // Init Language Label
            Object langLabelModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlFixedTextModel");

            XPropertySet xPSetLabel6 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, langLabelModel);

            xPSetLabel6.setPropertyValue("PositionX", new Integer(14));
            xPSetLabel6.setPropertyValue("PositionY", new Integer(93));
            xPSetLabel6.setPropertyValue("Width", new Integer(35));
            xPSetLabel6.setPropertyValue("Height", new Integer(9));
            xPSetLabel6.setPropertyValue("Name", _langLabelName);
            xPSetLabel6.setPropertyValue("Label", L10N_langLabelValue);
            //xPSetLabel6.setPropertyValue("MultiLine", new Boolean(true));

            // Init Bitrate Label
            Object bitrateLabelModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlFixedTextModel");

            XPropertySet xPSetLabel7 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, bitrateLabelModel);

            xPSetLabel7.setPropertyValue("PositionX", new Integer(14));
            xPSetLabel7.setPropertyValue("PositionY", new Integer(153));
            xPSetLabel7.setPropertyValue("Width", new Integer(35));
            xPSetLabel7.setPropertyValue("Height", new Integer(9));
            xPSetLabel7.setPropertyValue("Name", _bitrateLabelName);
            xPSetLabel7.setPropertyValue("Label", L10N_bitrateLabelValue);


            // Init UID Field
            Object uidFieldModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlEditModel");

            XPropertySet xPSetField1 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, uidFieldModel);

            xPSetField1.setPropertyValue("PositionX", new Integer(48));
            xPSetField1.setPropertyValue("PositionY", new Integer(15));
            xPSetField1.setPropertyValue("Width", new Integer(100));
            xPSetField1.setPropertyValue("Height", new Integer(14));
            xPSetField1.setPropertyValue("Name", _uidFieldName);
            xPSetField1.setPropertyValue("TabIndex", new Short((short) 0));
            //xPSetField1.setPropertyValue("Text", new String("UID"));


            // Init Doctitle Field            
            Object doctitleFieldModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlEditModel");

            XPropertySet xPSetField2 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, doctitleFieldModel);

            xPSetField2.setPropertyValue("PositionX", new Integer(48));
            xPSetField2.setPropertyValue("PositionY", new Integer(30));
            xPSetField2.setPropertyValue("Width", new Integer(100));
            xPSetField2.setPropertyValue("Height", new Integer(14));
            xPSetField2.setPropertyValue("Name", _doctitleFieldName);
            xPSetField2.setPropertyValue("TabIndex", new Short((short) 1));
            //xPSetField2.setPropertyValue("Text", new String("Doctitle"));


            // Init Creator Field
            Object creatorFieldModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlEditModel");

            XPropertySet xPSetField3 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, creatorFieldModel);

            xPSetField3.setPropertyValue("PositionX", new Integer(48));
            xPSetField3.setPropertyValue("PositionY", new Integer(45));
            xPSetField3.setPropertyValue("Width", new Integer(100));
            xPSetField3.setPropertyValue("Height", new Integer(14));
            xPSetField3.setPropertyValue("Name", _creatorFieldName);
            xPSetField3.setPropertyValue("TabIndex", new Short((short) 2));
            //xPSetField3.setPropertyValue("Text", new String("Creator"));


            // Init Publisher Field
            Object publisherFieldModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlEditModel");

            XPropertySet xPSetField4 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, publisherFieldModel);

            xPSetField4.setPropertyValue("PositionX", new Integer(48));
            xPSetField4.setPropertyValue("PositionY", new Integer(60));
            xPSetField4.setPropertyValue("Width", new Integer(100));
            xPSetField4.setPropertyValue("Height", new Integer(14));
            xPSetField4.setPropertyValue("Name", _publisherFieldName);
            xPSetField4.setPropertyValue("TabIndex", new Short((short) 3));
            //xPSetField4.setPropertyValue("Text", new String("Publisher"));


            // Init Producer Field
            Object producerFieldModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlEditModel");
            XPropertySet xPSetField9 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, producerFieldModel);

            xPSetField9.setPropertyValue("PositionX", new Integer(48));
            xPSetField9.setPropertyValue("PositionY", new Integer(75));
            xPSetField9.setPropertyValue("Width", new Integer(100));
            xPSetField9.setPropertyValue("Height", new Integer(14));
            xPSetField9.setPropertyValue("Name", _producerFieldName);
            xPSetField9.setPropertyValue("TabIndex", new Short((short) 4));
            //xPSetField9.setPropertyValue("Text", new String("Publisher"));


            // Init Language Field
            Object langFieldModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlEditModel");

            XPropertySet xPSetField5 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, langFieldModel);

            xPSetField5.setPropertyValue("ReadOnly", new Boolean(true));
            xPSetField5.setPropertyValue("Enabled", new Boolean(false));

            xPSetField5.setPropertyValue("PositionX", new Integer(48));
            xPSetField5.setPropertyValue("PositionY", new Integer(90));
            xPSetField5.setPropertyValue("Width", new Integer(19));
            xPSetField5.setPropertyValue("Height", new Integer(14));
            xPSetField5.setPropertyValue("Name", _langFieldName);
            xPSetField5.setPropertyValue("TabIndex", new Short((short) 5));
            //xPSetField5.setPropertyValue("MinTextLen",new Short((short)5));
            //xPSetField5.setPropertyValue("MaxTextLen",new Short((short)5));
            //xPSetField5.setPropertyValue("Text", new String("lang"));


            // Init Level ComboBox
            Object levelCBoxModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlCheckBoxModel");

            XPropertySet xPSetField6 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, levelCBoxModel);

            xPSetField6.setPropertyValue("PositionX", new Integer(48));
            xPSetField6.setPropertyValue("PositionY", new Integer(110));
            xPSetField6.setPropertyValue("Width", new Integer(100));
            xPSetField6.setPropertyValue("Height", new Integer(14));
            xPSetField6.setPropertyValue("Name", _levelCBoxName);
            xPSetField6.setPropertyValue("TabIndex", new Short((short) 6));
            xPSetField6.setPropertyValue("Label", L10N_levelLabelValue);
            //xPSetField6.setPropertyValue("State", new Short((short)0));
            //xPSetField6.setPropertyValue("TriState", new Boolean(true));


            // Init PageNumbering ComboBox
            Object pageCBoxModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlCheckBoxModel");

            XPropertySet xPSetField7 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, pageCBoxModel);

            xPSetField7.setPropertyValue("PositionX", new Integer(48));
            xPSetField7.setPropertyValue("PositionY", new Integer(120));
            xPSetField7.setPropertyValue("Width", new Integer(100));
            xPSetField7.setPropertyValue("Height", new Integer(14));
            xPSetField7.setPropertyValue("Name", _pageCBoxName);
            xPSetField7.setPropertyValue("TabIndex", new Short((short) 7));
            xPSetField7.setPropertyValue("Label", L10N_pageLabelValue);
            xPSetField7.setPropertyValue("State", new Short((short) 1));


            // Init CSS File ComboBox            
            Object cssCBoxModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlCheckBoxModel");

            XPropertySet xPSetField8 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, cssCBoxModel);

            xPSetField8.setPropertyValue("PositionX", new Integer(48));
            xPSetField8.setPropertyValue("PositionY", new Integer(130));
            xPSetField8.setPropertyValue("Width", new Integer(100));
            xPSetField8.setPropertyValue("Height", new Integer(14));
            xPSetField8.setPropertyValue("Name", _cssCBoxName);
            xPSetField8.setPropertyValue("TabIndex", new Short((short) 8));
            xPSetField8.setPropertyValue("Label", L10N_cssLabelValue);
            xPSetField8.setPropertyValue("State", new Short((short) 0));


            // Init Bitrate ComboBox
            bitrateControlModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlListBoxModel"); // replaces UnoControlComboBoxModel with UnoControlListBoxModel
            XPropertySet xPSetComboBox =
                    (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, bitrateControlModel);
            xPSetComboBox.setPropertyValue("PositionX", new Integer(48));
            xPSetComboBox.setPropertyValue("PositionY", new Integer(150));
            xPSetComboBox.setPropertyValue("Width", new Integer(45));
            xPSetComboBox.setPropertyValue("Height", new Integer(14));
            xPSetComboBox.setPropertyValue("Name", _bitrateControlName);
            xPSetComboBox.setPropertyValue("Dropdown", new Boolean(true));
            //xPSetComboBox.setPropertyValue("ReadOnly", new Boolean(true));
            xPSetComboBox.setPropertyValue("TabIndex", new Short((short) 9));
            xPSetComboBox.setPropertyValue("StringItemList",
                    new String[]{"32 kbit/s", "48 kbit/s", "64 kbit/s", "128 kbit/s"});
            //xPSetComboBox.setPropertyValue("Text", "32 kbits/s"); // from service UnoControlComboBoxModel
            short[] selectedItems = {(short) 0};
            xPSetComboBox.setPropertyValue("SelectedItems", selectedItems); // from service UnoControlListBoxModel

            // Init Fix Routine ComboBox
            Object fixRoutineCBoxModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlCheckBoxModel");

            XPropertySet xPSetField10 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, fixRoutineCBoxModel);

            xPSetField10.setPropertyValue("PositionX", new Integer(48));
            xPSetField10.setPropertyValue("PositionY", new Integer(165));
            xPSetField10.setPropertyValue("Width", new Integer(100));
            xPSetField10.setPropertyValue("Height", new Integer(14));
            xPSetField10.setPropertyValue("Name", _fixRoutineCBoxName);
            xPSetField10.setPropertyValue("TabIndex", new Short((short) 10));
            xPSetField10.setPropertyValue("Label", L10N_fixRoutineLabelValue);
            xPSetField10.setPropertyValue("State", new Short((short) 1));

            // Init Sent Detection ComboBox
            Object sentDetectionCBoxModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlCheckBoxModel");

            XPropertySet xPSetField11 = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, sentDetectionCBoxModel);

            xPSetField11.setPropertyValue("PositionX", new Integer(48));
            xPSetField11.setPropertyValue("PositionY", new Integer(175));
            xPSetField11.setPropertyValue("Width", new Integer(100));
            xPSetField11.setPropertyValue("Height", new Integer(14));
            xPSetField11.setPropertyValue("Name", _sentDetectionCBoxName);
            xPSetField11.setPropertyValue("TabIndex", new Short((short) 11));
            xPSetField11.setPropertyValue("Label", L10N_sentDetectionLabelValue);
            xPSetField11.setPropertyValue("State", new Short((short) 1));


            // posy buttons
            Integer posYButtons;
            if (isFullExport) {
                posYButtons = new Integer(190);
            } else {
                posYButtons = new Integer(150);
            }

            // Init Ok Button
            Object buttonModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlButtonModel");

            XPropertySet xPSetButton = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, buttonModel);

            xPSetButton.setPropertyValue("PositionX", new Integer(30));
            xPSetButton.setPropertyValue("PositionY", posYButtons);
            xPSetButton.setPropertyValue("Width", new Integer(50));
            xPSetButton.setPropertyValue("Height", new Integer(14));
            xPSetButton.setPropertyValue("Name", _okButtonName);
            xPSetButton.setPropertyValue("TabIndex", new Short((short) 12));
            xPSetButton.setPropertyValue("PushButtonType", (short) PushButtonType.OK_value);
            xPSetButton.setPropertyValue("Label", L10N_okButtonValue);

            // init Cancel Button
            Object cancelButtonModel = xMultiServiceFactory.createInstance(
                    "com.sun.star.awt.UnoControlButtonModel");
            XPropertySet xPSetCancelButton = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, cancelButtonModel);
            xPSetCancelButton.setPropertyValue("PositionX", new Integer(90));
            xPSetCancelButton.setPropertyValue("PositionY", posYButtons);
            xPSetCancelButton.setPropertyValue("Width", new Integer(50));
            xPSetCancelButton.setPropertyValue("Height", new Integer(14));
            xPSetCancelButton.setPropertyValue("Name", _cancelButtonName);
            xPSetCancelButton.setPropertyValue("TabIndex", new Short((short) 13));
            xPSetCancelButton.setPropertyValue("PushButtonType", (short) PushButtonType.CANCEL_value);
            xPSetCancelButton.setPropertyValue("Label", L10N_cancelButtonValue);


            // insert the control models into the dialog model
            xNameCont = (XNameContainer) UnoRuntime.queryInterface(
                    XNameContainer.class, dialogModel);
            xNameCont.insertByName(_okButtonName, buttonModel);
            xNameCont.insertByName(_cancelButtonName, cancelButtonModel);

            xNameCont.insertByName(_uidLabelName, uidLabelModel);
            xNameCont.insertByName(_doctitleLabelName, doctitleLabelModel);
            xNameCont.insertByName(_creatorLabelName, creatorLabelModel);
            xNameCont.insertByName(_publisherLabelName, publisherLabelModel);
            xNameCont.insertByName(_producerLabelName, producerLabelModel);
            xNameCont.insertByName(_langLabelName, langLabelModel);

            xNameCont.insertByName(_uidFieldName, uidFieldModel);
            xNameCont.insertByName(_doctitleFieldName, doctitleFieldModel);
            xNameCont.insertByName(_creatorFieldName, creatorFieldModel);
            xNameCont.insertByName(_publisherFieldName, publisherFieldModel);
            xNameCont.insertByName(_producerFieldName, producerFieldModel);
            xNameCont.insertByName(_langFieldName, langFieldModel);
            xNameCont.insertByName(_levelCBoxName, levelCBoxModel);
            xNameCont.insertByName(_pageCBoxName, pageCBoxModel);
            xNameCont.insertByName(_cssCBoxName, cssCBoxModel);

            if (isFullExport) {
                xNameCont.insertByName(_bitrateLabelName, bitrateLabelModel);
                xNameCont.insertByName(_bitrateControlName, bitrateControlModel);
                xNameCont.insertByName(_fixRoutineCBoxName, fixRoutineCBoxModel);
                xNameCont.insertByName(_sentDetectionCBoxName, sentDetectionCBoxModel);
            }

            // create the dialog control and set the model
            dialog = m_xMCF.createInstanceWithContext(
                    "com.sun.star.awt.UnoControlDialog", m_xContext);

            XControl xControl = (XControl) UnoRuntime.queryInterface(
                    XControl.class, dialog);

            XControlModel xControlModel = (XControlModel) UnoRuntime.queryInterface(
                    XControlModel.class, dialogModel);

            xControl.setModel(xControlModel);

            xControlCont = (XControlContainer) UnoRuntime.queryInterface(
                    XControlContainer.class, dialog);

            // create a peer
            Object toolkit = m_xMCF.createInstanceWithContext(
                    "com.sun.star.awt.Toolkit", m_xContext);

            XToolkit xToolkit = (XToolkit) UnoRuntime.queryInterface(XToolkit.class, toolkit);

            XWindow xWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, xControl);

            xWindow.setVisible(false);
            xControl.createPeer(xToolkit, null);

        } catch (com.sun.star.uno.Exception ex) {
            ex.printStackTrace();
        } catch (java.lang.Exception ex) {
            ex.printStackTrace();
        }
    }

    public boolean execute() {

        short ret;

        xDialog = (XDialog) UnoRuntime.queryInterface(XDialog.class, dialog);
        ret = xDialog.execute();

        updateValues();

        // dispose the dialog
        xComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, dialog);
        xComponent.dispose();

        if (ret == ((short) PushButtonType.OK_value)) {
            return true;
        } else {
            return false;
        }
    }

    boolean isPaginationEnable() {
        return isPageEnabled;
    }

    /**
     * Get values from export dialog and store them.
     */
    private void updateValues() {
        Object cbox;
        Object text;
        Object combobox;

        text = xControlCont.getControl(ExportDialog._uidFieldName);
        XTextComponent xText = (XTextComponent) UnoRuntime.queryInterface(
                XTextComponent.class, text);

        setUidValue(xText.getText());

        text = xControlCont.getControl(ExportDialog._doctitleFieldName);
        xText = (XTextComponent) UnoRuntime.queryInterface(
                XTextComponent.class, text);

        setDoctitleValue(xText.getText());

        text = xControlCont.getControl(ExportDialog._creatorFieldName);
        xText = (XTextComponent) UnoRuntime.queryInterface(
                XTextComponent.class, text);

        setCreatorValue(xText.getText());

        text = xControlCont.getControl(ExportDialog._publisherFieldName);
        xText = (XTextComponent) UnoRuntime.queryInterface(
                XTextComponent.class, text);

        setPublisherValue(xText.getText());

        text = xControlCont.getControl(ExportDialog._producerFieldName);
        xText = (XTextComponent) UnoRuntime.queryInterface(
                XTextComponent.class, text);

        setProducerValue(xText.getText());

        text = xControlCont.getControl(ExportDialog._langFieldName);
        xText = (XTextComponent) UnoRuntime.queryInterface(
                XTextComponent.class, text);

        setLangValue(xText.getText());

        cbox = xControlCont.getControl(ExportDialog._levelCBoxName);
        XCheckBox xCBox = (XCheckBox) UnoRuntime.queryInterface(
                XCheckBox.class, cbox);

        // 0 -> Unchecked
        // 1 -> Checked
        // 2 -> I Don't Know
        if (xCBox.getState() == 1) {
            setAlternateLevelMarkupValue(true);
        } else {
            setAlternateLevelMarkupValue(false);
        }
        cbox = xControlCont.getControl(ExportDialog._pageCBoxName);
        xCBox = (XCheckBox) UnoRuntime.queryInterface(
                XCheckBox.class, cbox);

        if (xCBox.getState() == 1) {
            setPageEnabled(true);
        } else {
            setPageEnabled(false);
        }

        cbox = xControlCont.getControl(ExportDialog._cssCBoxName);
        xCBox = (XCheckBox) UnoRuntime.queryInterface(
                XCheckBox.class, cbox);

        if (xCBox.getState() == 1) {
            setWriteCSS(true);
        } else {
            setWriteCSS(false);
        }

        if (isFullExport) {
            
            cbox = xControlCont.getControl(ExportDialog._fixRoutineCBoxName);
            xCBox = (XCheckBox) UnoRuntime.queryInterface(
                    XCheckBox.class, cbox);

            if (xCBox.getState() == 1) {
                setFixRoutine(true);
            } else {
                setFixRoutine(false);
            }

            cbox = xControlCont.getControl(ExportDialog._sentDetectionCBoxName);
            xCBox = (XCheckBox) UnoRuntime.queryInterface(
                    XCheckBox.class, cbox);

            if (xCBox.getState() == 1) {
                setSentDetection(true);
            } else {
                setSentDetection(false);
            }


            combobox = xControlCont.getControl(ExportDialog._bitrateControlName);
            XComboBox xComboBox = (XComboBox) UnoRuntime.queryInterface(
                    XComboBox.class, combobox);

            XPropertySet xPSetBitRate =
                    (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, bitrateControlModel);

            String bitrateString = null;
            String bitrateNumber = null;
            try {

                // bitrateString = xPSetComboBox.getPropertyValue("Text").toString(); // = code for combobox
                short[] selectedItems = (short[]) xPSetBitRate.getPropertyValue("SelectedItems");
                bitrateString = ( (String[]) xPSetBitRate.getPropertyValue("StringItemList") )[selectedItems[0]];
                bitrateNumber = bitrateString.substring(0, bitrateString.indexOf(" "));
                setBitrate(Integer.parseInt(bitrateString));

            } catch (IndexOutOfBoundsException e) {
                // bitrateString does not contain a space (e.g. "blah")
                // or is an empty string:
                System.err.println("IndexOutOfBoundsException" + 
                    "    bitrateString = " + bitrateString);
                setBitrate(32); // continue with default bitrate @@todo define private int
            } catch (NumberFormatException e) {
                // The first part of bitrateString is not a number:
                System.err.println("NumberFormatException" +
                    "    bitrateString = " + bitrateString);
                setBitrate(32); // continue with default bitrate
            } catch (Exception e) { // other exception
                // Only useful when running from inside NetBeans:
                System.err.println("Bitrate: exception other than\n" +
                    "    IndexOutOfBoundsException or NumberFormatException.");
                System.err.println("bitrateString = " + bitrateString);
                System.err.println(e.getMessage());
                System.err.println(e.getStackTrace());
            }

        }
    }

    public String getUid() {
        return uid;
    }

    public void setUid(String uid) {
        Object text = xControlCont.getControl(ExportDialog._uidFieldName);
        XTextComponent xText = (XTextComponent) UnoRuntime.queryInterface(
                XTextComponent.class, text);
        xText.setText(uid);

    }

    private void setUidValue(String uid) {
        this.uid = uid;
    }

    public String getDoctitle() {
        return doctitle;
    }

    public void setDoctitle(String doctitle) {
        Object text = xControlCont.getControl(ExportDialog._doctitleFieldName);
        XTextComponent xText = (XTextComponent) UnoRuntime.queryInterface(
                XTextComponent.class, text);
        xText.setText(doctitle);
    }

    private void setDoctitleValue(String doctitle) {
        this.doctitle = doctitle;
    }

    public String getCreator() {
        return creator;
    }

    public void setCreator(String creator) {
        Object text = xControlCont.getControl(ExportDialog._creatorFieldName);
        XTextComponent xText = (XTextComponent) UnoRuntime.queryInterface(
                XTextComponent.class, text);
        xText.setText(creator);
    }

    private void setCreatorValue(String creator) {
        this.creator = creator;
    }

    public String getPublisher() {
        return publisher;
    }

    public void setPublisher(String publisher) {
        Object text = xControlCont.getControl(ExportDialog._publisherFieldName);
        XTextComponent xText = (XTextComponent) UnoRuntime.queryInterface(
                XTextComponent.class, text);
        xText.setText(publisher);
    }

    private void setPublisherValue(String publisher) {
        this.publisher = publisher;
    }

    public String getProducer() {
        return producer;
    }

    public void setProducer(String producer) {
        Object text = xControlCont.getControl(ExportDialog._producerFieldName);
        XTextComponent xText = (XTextComponent) UnoRuntime.queryInterface(
                XTextComponent.class, text);
        xText.setText(producer);
    }

    private void setProducerValue(String producer) {
        this.producer = producer;
    }

    public String getLang() {
        return lang;
    }

    public void setLang(String lang) {
        Object text = xControlCont.getControl(ExportDialog._langFieldName);
        XTextComponent xText = (XTextComponent) UnoRuntime.queryInterface(
                XTextComponent.class, text);
        xText.setText(lang);
    }

    private void setLangValue(String lang) {
        this.lang = lang;
    }

    public boolean isAlternateLevelMarkup() {
        return alternateLevelMarkup;
    }

    public void setAlternateLevelMarkup(boolean alternateLevelMarkup) {
        Object cbox = xControlCont.getControl(ExportDialog._levelCBoxName);
        XCheckBox xCBox = (XCheckBox) UnoRuntime.queryInterface(
                XCheckBox.class, cbox);

        if (alternateLevelMarkup) {
            xCBox.setState((short) 1);
        } else {
            xCBox.setState((short) 0);
        }
    }

    private void setAlternateLevelMarkupValue(boolean alternateLevelMarkup) {
        this.alternateLevelMarkup = alternateLevelMarkup;
    }

    public boolean isPageEnabled() {
        return isPageEnabled;
    }

    public void setPageEnabled(boolean isPageEnabled) {
        this.isPageEnabled = isPageEnabled;
    }

    public boolean isWriteCSS() {
        return writeCSS;
    }

    public void setWriteCSS(boolean writeCSS) {
        this.writeCSS = writeCSS;
    }

    public boolean isFixRoutine() {
        return fixRoutine;
    }

    public void setFixRoutine(boolean fixRoutine) {
        this.fixRoutine = fixRoutine;
    }

    public boolean isSentDetection() {
        return sentDetection;
    }

    public void setSentDetection(boolean sentDetection) {
        this.sentDetection = sentDetection;
    }

    public int getBitrate() {
        return this.bitrate;
    }

    /**
     * Sets the bitrate, which should be one of 32, 48, 64 or 128 kbit/s.
     * If the bitrate is not one of these values, it is silently corrected to 32.
     * @param bitrate (as integer)
     */
    public void setBitrate(int bitrate) {
        if ( !(bitrate == 32 || bitrate == 48 || bitrate == 64 || bitrate == 128) ) {
            bitrate = 32; // 32 kbit/s is usually sufficient for TTS
        }
        this.bitrate = bitrate;
    }
}

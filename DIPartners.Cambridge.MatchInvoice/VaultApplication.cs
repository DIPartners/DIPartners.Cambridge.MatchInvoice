using System;
using System.Collections.Generic;
using System.Diagnostics;
using MFiles.VAF;
using MFiles.VAF.Common;
using MFiles.VAF.Configuration;
using MFiles.VAF.Core;
using MFilesAPI;

namespace DIPartners.Cambridge.MatchInvoice
{
    public class InvoiceDetail
    {
        public MFIdentifier PropertyID { get; set; }
        public MFConditionType ConditionType { get; set; }
        public String TypedValue { get; set; }

        public InvoiceDetail()
        {
            ConditionType = MFConditionType.MFConditionTypeEqual;
        }
    }
    /// <summary>
    /// The entry point for this Vault Application Framework application.
    /// </summary>
    /// <remarks>Examples and further information available on the developer portal: http://developer.m-files.com/. </remarks>
    public class VaultApplication
        : ConfigurableVaultApplicationBase<Configuration>
    {
        #region Set MFIdentifier
        [MFClass]
        public MFIdentifier Invoice_CD = "vClass.Invoice";
        [MFClass]
        public MFIdentifier InvoiceDetail_CD = "vClass.InvoiceDetail";
        [MFClass]
        public MFIdentifier PurchaseOrderDetail_CD = "vClass.PurchaseOrderDetail";
        [MFObjType]
        public MFIdentifier InvoiceDetail_OT = "vObject.InvoiceDetail";
        [MFObjType]
        public MFIdentifier PurchaseOrderDetail_OT = "vObject.PurchaseOrderDetail";

        [MFPropertyDef]
        public MFIdentifier Invoice_PD = "vProperty.Invoice"; 
        [MFPropertyDef]
        public MFIdentifier InvoiceName_PD = "vProperty.InvoiceName"; 
        [MFPropertyDef]
        public MFIdentifier InvoiceNumber_PD = "vProperty.InvoiceNumber";
        [MFPropertyDef]
        public MFIdentifier InvoiceDetailName_PD = "vProperty.InvoiceDetailName";
        [MFPropertyDef]
        public MFIdentifier PurchaseOrder_PD = "vProperty.PurchaseOrder";
        [MFPropertyDef]
        public MFIdentifier POReference_PD = "vProperty.POReference";

        [MFPropertyDef]
        public MFIdentifier ItemNumber_PD = "vProperty.ItemNumber";
        [MFPropertyDef]
        public MFIdentifier Quantity_PD = "vProperty.Quantity";
        [MFPropertyDef]
        public MFIdentifier UnitPrice_PD = "vProperty.UnitPrice";
        [MFPropertyDef]
        public MFIdentifier POLine_PD = "vProperty.POLine#";
        [MFPropertyDef]
        public MFIdentifier POItem_PD = "vProperty.POItem";
        [MFPropertyDef]
        public MFIdentifier OrderedQty_PD = "vProperty.OrderedQty";

        [MFPropertyDef]
        public MFIdentifier POLineExtension_PD = "vProperty.POLineExtension";
        [MFPropertyDef]
        public MFIdentifier InvoiceLineExtension_PD = "vProperty.InvoiceLineExtension";
        [MFPropertyDef]
        public MFIdentifier InvoiceLineNumber_PD = "vProperty.InvoiceLineNumber";
        [MFPropertyDef]
        public MFIdentifier PurchaseOrderDetail_PD = "vProperty.PurchaseOrderDetail";
        [MFPropertyDef]
        public MFIdentifier PurchaseOrderDetailName_PD = "vProperty.PurchaseOrderDetailName";

        #endregion
        // Event Handler Before Create New Invoice Finalize
        [EventHandler(MFEventHandlerType.MFEventHandlerAfterCheckInChangesFinalize, Class = "vClass.Invoice")]
        public void CreateNewInvoice(EventHandlerEnvironment env)
        {
            var Vault = env.ObjVerEx.Vault;
            var oCurrObjVals = Vault.ObjectPropertyOperations.GetProperties(env.ObjVerEx.ObjVer);

            // Search Current Invoice Number
            var InvoiceNum = SearchPropertyValue(oCurrObjVals, InvoiceNumber_PD);
            if (InvoiceNum == "") return;

            var POReference = SearchPropertyValue(oCurrObjVals, POReference_PD);
            if (POReference == "") return;

            // Get Invoice Data to find PO number
            List<ObjVerEx> objInvoices = FindObjects(Vault, InvoiceDetail_CD, Invoice_PD, MFDataType.MFDatatypeLookup, env.ObjVerEx.ObjVer.ID.ToString());
            List<ObjVerEx> ojbPOs = FindObjects(Vault, PurchaseOrderDetail_CD, PurchaseOrder_PD, MFDataType.MFDatatypeText, POReference);

            double Invext = 0d;
            double POext = 0d;
            foreach (var PO in ojbPOs)
            {
                POext += Convert.ToDouble(Vault.ObjectPropertyOperations.GetProperties(PO.ObjVer)
                                .SearchForProperty(POLineExtension_PD).TypedValue.DisplayValue);
            }

            if (objInvoices == null)
            {
                if (Invext != POext)
                {
                    foreach (var ojbPO in ojbPOs)
                    {
                        CreateNewDetails(env.ObjVerEx, ojbPO);
                    }
                }
            }
            else
            {
                foreach (var invoice in objInvoices)
                {
                    Invext += Convert.ToDouble(Vault.ObjectPropertyOperations.GetProperties(invoice.ObjVer)
                                    .SearchForProperty(InvoiceLineExtension_PD).TypedValue.DisplayValue);
                }

                if (objInvoices.Count < ojbPOs.Count || Invext != POext)
                {
                    foreach (var ojbPO in ojbPOs)
                    {
                        bool isFoundInvoice = false;
                        foreach (var invItem in objInvoices)
                        {
                            #region What's gonna be in PurchaseOrderDetail
                            /*if (Vault.ObjectPropertyOperations.GetProperties(ojbPO.ObjVer).SearchForProperty(PurchaseOrderDetailName_PD) ==
                            Vault.ObjectPropertyOperations.GetProperties(invItem.ObjVer).SearchForProperty(PurchaseOrderDetail_PD))
                            {
                                isFoundInvoice = true;
                                break;
                            }*/
                            #endregion

                            if (Vault.ObjectPropertyOperations.GetProperties(ojbPO.ObjVer).SearchForProperty(POLine_PD).TypedValue.DisplayValue ==
                            Vault.ObjectPropertyOperations.GetProperties(invItem.ObjVer).SearchForProperty(InvoiceLineNumber_PD).TypedValue.DisplayValue)
                            {
                                isFoundInvoice = true;
                                break;
                            }
                        }
                        if (!isFoundInvoice)
                        {
                            CreateNewDetails(env.ObjVerEx, ojbPO);
                        }
                    }
                }
            }            
        }

        public void CreateNewDetails(ObjVerEx objVer, ObjVerEx ojbPO)
        {
            var propertyValues = new PropertyValues();

            //set class
            var classPropertyValue = new PropertyValue()
            {
                PropertyDef = (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass
            };
            classPropertyValue.Value.SetValue(MFDataType.MFDatatypeLookup, objVer.Vault.ClassOperations.GetObjectClass(InvoiceDetail_CD).ID);
            propertyValues.Add(-1, classPropertyValue);

            // set Name or Title
            var TitleProperties = objVer.Vault.ObjectPropertyOperations.GetProperties(objVer.ObjVer, true);
            var propTitle = TitleProperties.SearchForProperty((int)MFBuiltInPropertyDef.MFBuiltInPropertyDefNameOrTitle);

            var nameOrTitlePropertyValue = new PropertyValue()
            {
                PropertyDef = (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefNameOrTitle
            };
            nameOrTitlePropertyValue.Value.SetValue(propTitle.TypedValue.DataType, propTitle.TypedValue.DisplayValue);
            propertyValues.Add(-1, nameOrTitlePropertyValue);

            // set Invoice
            var NewInvoiceLookup = new Lookup();
            NewInvoiceLookup.ObjectType = objVer.ObjVer.Type;
            NewInvoiceLookup.Item = objVer.ObjVer.ID;
            NewInvoiceLookup.DisplayValue = TitleProperties.SearchForProperty(InvoiceName_PD).TypedValue.DisplayValue;
            var newInvoice = new PropertyValue()
            {
                PropertyDef = 1058  //(int)MFBuiltInPropertyDef.MFBuiltInPropertyDefNameOrTitle
            };
            newInvoice.Value.SetValue(MFDataType.MFDatatypeLookup, NewInvoiceLookup);
            propertyValues.Add(-1, newInvoice);


            PropertyValues Inv = objVer.Vault.ObjectPropertyOperations.GetProperties(objVer.ObjVer);
            PropertyValues PO = objVer.Vault.ObjectPropertyOperations.GetProperties(ojbPO.ObjVer);
            propertyValues.Add(-1, GetPropertyValue(PO, POLine_PD, Inv, InvoiceLineNumber_PD));
            propertyValues.Add(-1, GetPropertyValue(PO, POItem_PD, Inv, ItemNumber_PD));
            propertyValues.Add(-1, GetPropertyValue(PO, OrderedQty_PD, Inv, Quantity_PD));
            propertyValues.Add(-1, GetPropertyValue(PO, UnitPrice_PD, Inv, UnitPrice_PD));
            propertyValues.Add(-1, GetPropertyValue(PO, POLineExtension_PD, Inv, InvoiceLineExtension_PD));
            propertyValues.Add(-1, GetPropertyValue(PO, PurchaseOrder_PD, Inv, PurchaseOrderDetail_PD));

            ObjectVersionAndProperties ppts = objVer.Vault.ObjectOperations.CreateNewObject(InvoiceDetail_OT, propertyValues);

            objVer.Vault.ObjectOperations.CheckIn(ppts.ObjVer);
        }

        public PropertyValue GetPropertyValue(PropertyValues POPpvs, MFIdentifier PropertyDef, PropertyValues InvPpvs, MFIdentifier SetDef = null)
        {
            var ppValue = new PropertyValue();
            ppValue.PropertyDef = SetDef;
            if (SetDef == null) SetDef = PropertyDef;

            var POPpt = POPpvs.SearchForProperty(PropertyDef);

            PropertyValue InvPpt = (PropertyDef == POLine_PD || PropertyDef == PurchaseOrder_PD) ? POPpt : InvPpvs.SearchForProperty(SetDef);

            string strVal =  SearchPropertyValue(POPpvs, SetDef, POPpt);

            if (SetDef == ItemNumber_PD) {
                string[] displayValues = strVal.Split('=');
                strVal = displayValues[0];
            }

            MFDataType dataType = (PropertyDef == POLine_PD) ? MFDataType.MFDatatypeInteger : InvPpt.TypedValue.DataType;
            ppValue.Value.SetValue(dataType, strVal);

            return ppValue;
        }

        public List<ObjVerEx> FindObjects(Vault vault, MFIdentifier ClassAlias, MFIdentifier PDAlias, MFDataType PDType, String findValue)
        {
            // Create our search builder.
            var searchBuilder = new MFSearchBuilder(vault);
            // Add an object type filter.
            searchBuilder.Class(ClassAlias);
            // Add a "not deleted" filter.
            searchBuilder.Deleted(false);
            List<ObjVerEx> searchResults;
            searchBuilder.Property(PDAlias, PDType, findValue);

            searchResults = searchBuilder.FindEx();

            return (searchResults.Count != 0) ? searchResults : null;
        }
        public string SearchPropertyValue(PropertyValues ppvs, MFIdentifier def, PropertyValue defaultPpt = null)
        {
            var ppt = defaultPpt;
            if (ppt == null) ppt = ppvs.SearchForProperty(def);

            return (ppt.TypedValue.DataType == MFDataType.MFDatatypeLookup) ?
                        ppt.TypedValue.GetLookupID().ToString() : ppt.TypedValue.DisplayValue;
        }
    }
}
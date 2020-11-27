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
        [EventHandler(MFEventHandlerType.MFEventHandlerAfterCheckInChangesFinalize, Class = "vClass.InvoiceDetail")]
        public void CreateNewInvoice(EventHandlerEnvironment env)
        {
            var Vault = env.ObjVerEx.Vault;
            var oCurrObjVals = Vault.ObjectPropertyOperations.GetProperties(env.ObjVerEx.ObjVer, true);
            var InvoiceProperty = oCurrObjVals.SearchForProperty(Invoice_PD).TypedValue.GetValueAsLookup();

            var InvoiceObjID = new ObjID();
            InvoiceObjID.SetIDs(InvoiceProperty.ObjectType, InvoiceProperty.Item);
            var InvoiceObjVer = Vault.ObjectOperations.GetLatestObjVer(InvoiceObjID, false, true);
            var InvoiceoObjProperties = Vault.ObjectOperations.GetObjectVersionAndProperties(InvoiceObjVer);
            var POReference = InvoiceoObjProperties.Properties.SearchForProperty(POReference_PD).TypedValue.Value.ToString();
            if (POReference == "") return;

            // Get Data
            List<ObjVerEx> objInvoices = FindObjects(Vault, InvoiceDetail_CD, Invoice_PD, MFDataType.MFDatatypeLookup, env.ObjVerEx.ObjVer.ID.ToString());
            List<ObjVerEx> objPOs = FindObjects(Vault, PurchaseOrderDetail_CD, PurchaseOrder_PD, MFDataType.MFDatatypeText, POReference);

            double Invext = 0d;
            double POext = 0d;
            foreach (var PO in objPOs)
            {
                POext += Convert.ToDouble(Vault.ObjectPropertyOperations.GetProperties(PO.ObjVer)
                                .SearchForProperty(POLineExtension_PD).TypedValue.DisplayValue);
            }

            if (objInvoices == null)
            {
                if (Invext != POext)
                {
                    foreach (var objPO in objPOs)
                    {
                        CreateNewDetails(env.ObjVerEx, objPO);
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

                if (objInvoices.Count < objPOs.Count || Invext != POext)
                {
                    foreach (var objPO in objPOs)
                    {
                        bool isFoundInvoice = false;
                        foreach (var invItem in objInvoices)
                        {
                            if (Vault.ObjectPropertyOperations.GetProperties(objPO.ObjVer).SearchForProperty(PurchaseOrderDetailName_PD).TypedValue.DisplayValue ==
                            Vault.ObjectPropertyOperations.GetProperties(invItem.ObjVer).SearchForProperty(PurchaseOrderDetail_PD).TypedValue.DisplayValue)
                            {
                                isFoundInvoice = true;
                                break;
                            }
                        }
                        if (!isFoundInvoice)
                        {
                            CreateNewDetails(env.ObjVerEx, objPO);
                        }
                    }
                }
            }            
        }

        public void CreateNewDetails(ObjVerEx objVer, ObjVerEx objPO)
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
            var NewInvoiceLookup = new Lookup()
            {
                ObjectType = objVer.ObjVer.Type,
                Item = objVer.ObjVer.ID,
                DisplayValue = TitleProperties.SearchForProperty(InvoiceName_PD).TypedValue.DisplayValue
            };
            var newInvoice = new PropertyValue()
            {
                PropertyDef = Invoice_PD.ID      //1058
            };
            newInvoice.Value.SetValue(MFDataType.MFDatatypeLookup, NewInvoiceLookup);
            propertyValues.Add(-1, newInvoice);

            // set PODetail
            var NewPOLookup = new Lookup()
            {
                ObjectType = objPO.ObjVer.Type,
                Item = objPO.ObjVer.ID,
                DisplayValue = objPO.Title
            };
            var newPO = new PropertyValue()
            {
                PropertyDef = PurchaseOrderDetail_PD.ID      //1177
            };
            newPO.Value.SetValue(MFDataType.MFDatatypeLookup, NewPOLookup);
            propertyValues.Add(-1, newPO);

            PropertyValues PO = objVer.Vault.ObjectPropertyOperations.GetProperties(objPO.ObjVer);
            propertyValues.Add(-1, GetPropertyValue(objVer.Vault, PO, POLine_PD, InvoiceLineNumber_PD));
            propertyValues.Add(-1, GetPropertyValue(objVer.Vault, PO, POItem_PD, ItemNumber_PD));
            propertyValues.Add(-1, GetPropertyValue(objVer.Vault, PO, OrderedQty_PD, Quantity_PD));
            propertyValues.Add(-1, GetPropertyValue(objVer.Vault, PO, UnitPrice_PD, UnitPrice_PD));
            propertyValues.Add(-1, GetPropertyValue(objVer.Vault, PO, POLineExtension_PD, InvoiceLineExtension_PD));

            ObjectVersionAndProperties ppts = objVer.Vault.ObjectOperations.CreateNewObject(InvoiceDetail_OT, propertyValues);
            
            objVer.Vault.ObjectOperations.CheckIn(ppts.ObjVer);

        }

        public PropertyValue GetPropertyValue(Vault vault, PropertyValues POPpvs, MFIdentifier PODef, MFIdentifier InvDef)
        {
            var ppValue = new PropertyValue();
            ppValue.PropertyDef = InvDef.ID;

            string strVal = GetPropertyValue(POPpvs.SearchForProperty(PODef));
            if (InvDef == ItemNumber_PD)
            {
                string[] displayValues = strVal.Split('=');
                strVal = displayValues[0];
            }
            ppValue.Value.SetValue(vault.PropertyDefOperations.GetPropertyDef(InvDef.ID).DataType, strVal);

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

        public string GetPropertyValue(PropertyValue POPty)
        {
            return (POPty.TypedValue.DataType == MFDataType.MFDatatypeLookup) ?
                        POPty.TypedValue.GetLookupID().ToString() : POPty.TypedValue.DisplayValue;
        }
    }
}
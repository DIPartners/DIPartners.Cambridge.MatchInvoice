using System;
using System.Diagnostics;
using MFiles.VAF;
using MFiles.VAF.Common;
using MFiles.VAF.Configuration;
using MFiles.VAF.Core;
using MFilesAPI;

namespace DIPartners.Cambridge.MatchInvoice
{
    /// <summary>
    /// The entry point for this Vault Application Framework application.
    /// </summary>
    /// <remarks>Examples and further information available on the developer portal: http://developer.m-files.com/. </remarks>
    public class VaultApplication
        : ConfigurableVaultApplicationBase<Configuration>
    {
        #region Set MFIdentifier

        #endregion

        [StateAction("vState.InvoiceReview")]
        public void CreateNewInvoice(StateEnvironment env)
        {
            var Vault = env.ObjVerEx.Vault;
        }
    }
}
﻿//------------------------------------------------------------------------------
// <auto-generated>
//     這段程式碼是由工具產生的。
//     執行階段版本:4.0.30319.34003
//
//     對這個檔案所做的變更可能會造成錯誤的行為，而且如果重新產生程式碼，
//     變更將會遺失。
// </auto-generated>
//------------------------------------------------------------------------------

namespace UpdateRecordReport.Properties {
    using System;
    
    
    /// <summary>
    ///   用於查詢當地語系化字串等的強類型資源類別。
    /// </summary>
    // 這個類別是自動產生的，是利用 StronglyTypedResourceBuilder
    // 類別透過 ResGen 或 Visual Studio 這類工具。
    // 若要加入或移除成員，請編輯您的 .ResX 檔，然後重新執行 ResGen
    // (利用 /str 選項)，或重建您的 VS 專案。
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   傳回這個類別使用的快取的 ResourceManager 執行個體。
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("UpdateRecordReport.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   覆寫目前執行緒的 CurrentUICulture 屬性，對象是所有
        ///   使用這個強類型資源類別的資源查閱。
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        /// </summary>
        internal static System.Drawing.Bitmap Report {
            get {
                object obj = ResourceManager.GetObject("Report", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   查詢類型 System.Byte[] 的當地語系化資源。
        /// </summary>
        internal static byte[] Template {
            get {
                object obj = ResourceManager.GetObject("Template", resourceCulture);
                return ((byte[])(obj));
            }
        }
        
        /// <summary>
        ///   查詢類似 &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;
        ///&lt;異動代號對照表&gt;
        ///  &lt;異動&gt;
        ///    &lt;代號&gt;001&lt;/代號&gt;
        ///    &lt;原因及事項&gt;持國民中學畢業證明書者(含國中補校)&lt;/原因及事項&gt;
        ///    &lt;分類&gt;新生異動&lt;/分類&gt;
        ///  &lt;/異動&gt;
        ///  &lt;異動&gt;
        ///    &lt;代號&gt;002&lt;/代號&gt;
        ///    &lt;原因及事項&gt;持國民中學補習學校資格證明書者&lt;/原因及事項&gt;
        ///    &lt;分類&gt;新生異動&lt;/分類&gt;
        ///  &lt;/異動&gt;
        ///  &lt;異動&gt;
        ///    &lt;代號&gt;003&lt;/代號&gt;
        ///    &lt;原因及事項&gt;持國民中學補習學校結業證明書者&lt;/原因及事項&gt;
        ///    &lt;分類&gt;新生異動&lt;/分類&gt;
        ///  &lt;/異動&gt;
        ///  &lt;異動&gt;
        ///    &lt;代號&gt;004&lt;/代號&gt;
        ///    &lt;原因及事項&gt;持國民中學修(結)業證明書者(修畢三學年全部課程)&lt;/原因及事項&gt;
        ///    &lt;分類&gt;新生異動&lt;/分類&gt;
        ///  &lt;/異動&gt;
        ///  &lt;異動&gt;
        ///    &lt;代號&gt;005&lt;/代號&gt;
        ///    &lt;原因及事項&gt;持國民中學畢業程度學力鑑定考試及格證明書者&lt;/原因及事項&gt;
        ///    &lt;分類&gt;新生異動&lt;/分類&gt;
        ///  [字串的其餘部分已遭截斷]&quot;; 的當地語系化字串。
        /// </summary>
        internal static string UpdateCode_SHD {
            get {
                return ResourceManager.GetString("UpdateCode_SHD", resourceCulture);
            }
        }
    }
}

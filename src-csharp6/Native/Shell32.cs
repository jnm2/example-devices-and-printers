using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using FILETIME = System.Runtime.InteropServices.ComTypes.FILETIME;

// ReSharper disable InconsistentNaming
#pragma warning disable 649

namespace Example.Native
{
    public static class Shell32
    {
        public const short CF_UNICODETEXT = 13;

        public static class FOLDERID
        {
            public static readonly Guid ControlPanelFolder = new Guid("82A74AEB-AEB4-465C-A014-D097EE346D63");
            public static readonly Guid PrintersFolder = new Guid("76FC4E2D-D6AD-4519-A663-37BD56068185");
        }

        public static class PKEY
        {
            public static readonly PROPERTYKEY Devices_FriendlyName = new PROPERTYKEY(new Guid("656A3BB3-ECC0-43FD-8477-4AE0404A96CD"), 12288);
            public static readonly PROPERTYKEY Devices_CategoryIds = new PROPERTYKEY(new Guid("78C34FC8-104A-4ACA-9EA4-524D52996E57"), 90);
            public static readonly PROPERTYKEY Devices_Status1 = new PROPERTYKEY(new Guid("78C34FC8-104A-4ACA-9EA4-524D52996E57"), 81);
            public static readonly PROPERTYKEY Devices_Status2 = new PROPERTYKEY(new Guid("78C34FC8-104A-4ACA-9EA4-524D52996E57"), 82);
        }

        public class CoTaskMemSafeHandle : SafeHandle
        {
            protected CoTaskMemSafeHandle() : base(IntPtr.Zero, true)
            {
            }

            public override bool IsInvalid => handle == IntPtr.Zero;

            protected override bool ReleaseHandle()
            {
                Marshal.FreeCoTaskMem(handle);
                return true;
            }
        }

        public sealed class ItemIdListSafeHandle : CoTaskMemSafeHandle
        {
            private ItemIdListSafeHandle()
            {
            }
        }

        public sealed class CoTaskStringSafeHandle : CoTaskMemSafeHandle
        {
            private CoTaskStringSafeHandle()
            {
            }

            public static CoTaskStringSafeHandle FromPointer(IntPtr pointer)
            {
                var r = new CoTaskStringSafeHandle();
                r.SetHandle(pointer);
                return r;
            }

            public string ReadAndFree()
            {
                using (this)
                    return Marshal.PtrToStringUni(handle);
            }
        }


        [DllImport("shell32.dll")]
        public static extern HResult SHGetKnownFolderIDList([In, MarshalAs(UnmanagedType.LPStruct)] Guid rfid, KF_FLAG dwFlags, IntPtr hToken, out ItemIdListSafeHandle ppidl);

        [Flags]
        public enum KF_FLAG : uint
        {
            DEFAULT = 0x00000000,
            SIMPLE_IDLIST = 0x00000100,
            NOT_PARENT_RELATIVE = 0x00000200,
            DEFAULT_PATH = 0x00000400,
            INIT = 0x00000800,
            NO_ALIAS = 0x00001000,
            DONT_UNEXPAND = 0x00002000,
            DONT_VERIFY = 0x00004000,
            CREATE = 0x00008000,
            NO_APPCONTAINER_REDIRECTION = 0x00010000,
            FORCE_APPCONTAINER_REDIRECTION = 0x00020000,
            ALIAS_ONLY = 0x80000000
        }

        [DllImport("shell32.dll")]
        public static extern HResult SHBindToObject(IntPtr psf, ItemIdListSafeHandle pidl, IBindCtx pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, [MarshalAs(UnmanagedType.IUnknown, IidParameterIndex = 3)] out object ppv);

        [DllImport("shell32.dll")]
        public static extern ItemIdListSafeHandle ILCombine(ItemIdListSafeHandle pIDLParent, ItemIdListSafeHandle pIDLChild);

        [DllImport("shell32.dll")]
        public static extern HResult SHCreateItemFromIDList(ItemIdListSafeHandle pidl, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, [MarshalAs(UnmanagedType.IUnknown, IidParameterIndex = 1)] out object ppv);



        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
        public interface IShellItem
        {
            [return: MarshalAs(UnmanagedType.IUnknown, IidParameterIndex = 2)]
            object BindToHandler(IBindCtx pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid rbhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid);

            IShellItem GetParent();

            IntPtr GetDisplayName(SIGDN sigdnName);

            SFGAO GetAttributes(SFGAO sfgaoMask);

            int Compare(IShellItem psi, SICHINTF hint);
        }

        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("7e9fb0d3-919f-4307-ab2e-9b1860310c93")]
        public interface IShellItem2
        {
            [return: MarshalAs(UnmanagedType.IUnknown, IidParameterIndex = 2)]
            object BindToHandler(IBindCtx pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid rbhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid);

            IShellItem GetParent();

            CoTaskStringSafeHandle GetDisplayName(SIGDN sigdnName);

            SFGAO GetAttributes(SFGAO sfgaoMask);

            int Compare(IShellItem psi, SICHINTF hint);

            [return: MarshalAs(UnmanagedType.IUnknown, IidParameterIndex = 1)]
            object GetPropertyStore(GPS flags, [MarshalAs(UnmanagedType.LPStruct)] Guid riid);

            [return: MarshalAs(UnmanagedType.IUnknown, IidParameterIndex = 2)]
            object GetPropertyStoreWithCreateObject(GPS flags, object punkCreateObject, [MarshalAs(UnmanagedType.LPStruct)] Guid riid);

            [return: MarshalAs(UnmanagedType.IUnknown, IidParameterIndex = 3)]
            object GetPropertyStoreForKeys([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] PROPERTYKEY[] rgKeys, uint cKeys, GPS flags, [MarshalAs(UnmanagedType.LPStruct)] Guid riid);

            [return: MarshalAs(UnmanagedType.IUnknown, IidParameterIndex = 1)]
            object GetPropertyDescriptionList([In] ref PROPERTYKEY keyType, [MarshalAs(UnmanagedType.LPStruct)] Guid riid);

            void Update(IBindCtx pbc);

            [PreserveSig]
            HResult GetProperty([In] ref PROPERTYKEY key, PropVariantSafeHandle ppropvar);

            [PreserveSig]
            HResult GetCLSID([In] ref PROPERTYKEY key, out Guid pclsid);

            [PreserveSig]
            HResult GetFileTime([In] ref PROPERTYKEY key, out FILETIME pft);

            [PreserveSig]
            HResult GetInt32([In] ref PROPERTYKEY key, out int pi);

            [PreserveSig]
            HResult GetString([In] ref PROPERTYKEY key, out CoTaskStringSafeHandle ppsz);

            [PreserveSig]
            HResult GetUInt32([In] ref PROPERTYKEY key, out uint pui);

            [PreserveSig]
            HResult GetUInt64([In] ref PROPERTYKEY key, out ulong pull);

            [PreserveSig]
            HResult GetBool([In] ref PROPERTYKEY key, out bool pf);
        }

        public sealed class PropVariantSafeHandle : CoTaskMemSafeHandle
        {
            public PropVariantSafeHandle()
            {
                SetHandle(Marshal.AllocCoTaskMem(8 + (IntPtr.Size * 2)));
            }

            public override bool IsInvalid => handle == IntPtr.Zero;

            protected override bool ReleaseHandle()
            {
                FreePropVariantArray(1, handle);
                return base.ReleaseHandle();
            }

            [DllImport("ole32.dll")]
            private static extern HResult FreePropVariantArray(uint cVariants, IntPtr rgvars);

            [DllImport("propsys.dll")]
            private static extern HResult PropVariantToStringVectorAlloc(PropVariantSafeHandle propvar, out CoTaskMemSafeHandle pprgsz, out uint pcElem);

            public unsafe VarEnum VarType => *(VarEnum*)handle;

            private unsafe struct CALPWSTR
            {
                public readonly uint cElems;
                public readonly char** pElems;
            }

            public unsafe string[] ToStringVector()
            {
                if (IsInvalid) throw new InvalidOperationException("The handle is invalid.");

                if (VarType == (VarEnum.VT_VECTOR | VarEnum.VT_LPWSTR))
                {
                    var vector = (CALPWSTR*)(handle + 8);
                    var cElems = vector->cElems;
                    if (cElems == 0) return Array.Empty<string>();

                    var r = new string[cElems];
                    var currentElementPointer = *vector->pElems;

                    for (var i = 0; i < r.Length; i++)
                    {
                        r[i] = new string(currentElementPointer);
                        currentElementPointer++;
                    }

                    return r;
                }
                else
                {
                    CoTaskMemSafeHandle buffer;
                    uint cElems;
                    if (PropVariantToStringVectorAlloc(this, out buffer, out cElems).Code != WinErrorCode.InvalidParameter)
                    {
                        using (buffer)
                        {
                            if (cElems == 0) return Array.Empty<string>();

                            var mustRelease = true;
                            buffer.DangerousAddRef(ref mustRelease);
                            try
                            {
                                var r = new string[cElems];
                                var currentElementPointer = (IntPtr*) buffer.DangerousGetHandle();

                                for (var i = 0; i < r.Length; i++)
                                {
                                    r[i] = CoTaskStringSafeHandle.FromPointer(*currentElementPointer).ReadAndFree();
                                    currentElementPointer++;
                                }

                                return r;
                            }
                            finally
                            {
                                if (mustRelease) buffer.DangerousRelease();
                            }
                        }
                    }
                }

                throw new NotImplementedException(VarType.ToString());
            }
        }

        public struct PROPERTYKEY
        {
            public readonly Guid fmtid;
            public readonly uint pid;

            public PROPERTYKEY(Guid fmtid, uint pid)
            {
                this.fmtid = fmtid;
                this.pid = pid;
            }
        }

        public enum GPS
        {
            DEFAULT = 0x00000000,
            HANDLERPROPERTIESONLY = 0x00000001,
            READWRITE = 0x00000002,
            TEMPORARY = 0x00000004,
            FASTPROPERTIESONLY = 0x00000008,
            OPENSLOWITEM = 0x00000010,
            DELAYCREATION = 0x00000020,
            BESTEFFORT = 0x00000040,
            NO_OPLOCK = 0x00000080,
            PREFERQUERYPROPERTIES = 0x00000100,
            EXTRINSICPROPERTIES = 0x00000200,
            EXTRINSICPROPERTIESONLY = 0x00000400,
            VOLATILEPROPERTIES = 0x00000800,
            VOLATILEPROPERTIESONLY = 0x00001000
        }

        [Flags]
        public enum SIGDN : uint
        {
            NORMALDISPLAY = 0x00000000,
            PARENTRELATIVEPARSING = 0x80018001,
            DESKTOPABSOLUTEPARSING = 0x80028000,
            PARENTRELATIVEEDITING = 0x80031001,
            DESKTOPABSOLUTEEDITING = 0x8004c000,
            FILESYSPATH = 0x80058000,
            URL = 0x80068000,
            PARENTRELATIVEFORADDRESSBAR = 0x8007c001,
            PARENTRELATIVE = 0x80080001,
            PARENTRELATIVEFORUI = 0x80094001
        }

        [Flags]
        public enum SICHINTF : uint
        {
            SICHINT_DISPLAY = 0x00000000,
            SICHINT_ALLFIELDS = 0x80000000,
            SICHINT_CANONICAL = 0x10000000,
            SICHINT_TEST_FILESYSPATH_IF_NOT_EQUAL = 0x20000000
        }



        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
        public interface IShellItemImageFactory
        {
            Gdi32.BitmapSafeHandle GetImage(POINT size, SIIGBF flags);
        }

        [Flags]
        public enum SIIGBF
        {
            SIIGBF_RESIZETOFIT = 0x00000000,
            SIIGBF_BIGGERSIZEOK = 0x00000001,
            SIIGBF_MEMORYONLY = 0x00000002,
            SIIGBF_ICONONLY = 0x00000004,
            SIIGBF_THUMBNAILONLY = 0x00000008,
            SIIGBF_INCACHEONLY = 0x00000010,
            SIIGBF_CROPTOSQUARE = 0x00000020,
            SIIGBF_WIDETHUMBNAILS = 0x00000040,
            SIIGBF_ICONBACKGROUND = 0x00000080,
            SIIGBF_SCALEUP = 0x00000100
        }



        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214E6-0000-0000-C000-000000000046")]
        public interface IShellFolder
        {
            void ParseDisplayName(IntPtr hwnd, IBindCtx pbc, string pszDisplayName, IntPtr pchEaten, out ItemIdListSafeHandle ppidl, IntPtr pdwAttributes);

            IEnumIDList EnumObjects(IntPtr hwnd, SHCONTF grfFlags);

            [return: MarshalAs(UnmanagedType.IUnknown, IidParameterIndex = 2)]
            object BindToObject(ItemIdListSafeHandle pidl, IBindCtx pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid riid);

            [return: MarshalAs(UnmanagedType.IUnknown, IidParameterIndex = 2)]
            object BindToStorage(ItemIdListSafeHandle pidl, IBindCtx pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid riid);

            [PreserveSig]
            HResult CompareIDs(IntPtr lParam, ItemIdListSafeHandle pidl1, ItemIdListSafeHandle pidl2);

            [return: MarshalAs(UnmanagedType.IUnknown, IidParameterIndex = 1)]
            object CreateViewObject(IntPtr hwndOwner, [MarshalAs(UnmanagedType.LPStruct)] Guid riid);

            void GetAttributesOf(uint cidl, SafeHandleArray apidl, [In, Out] ref SFGAO rgfInOut);

            [return: MarshalAs(UnmanagedType.IUnknown, IidParameterIndex = 3)]
            object GetUIObjectOf(IntPtr hwndOwner, uint cidl, SafeHandleArray apidl, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, IntPtr rgfReserved);

            IntPtr GetDisplayNameOf(ItemIdListSafeHandle pidl, SHGDN uFlags);

            IntPtr SetNameOf(IntPtr hwnd, ItemIdListSafeHandle pidl, [MarshalAs(UnmanagedType.LPWStr)] string pszName, SHGDN uFlags);
        }

        [Flags]
        public enum SHCONTF
        {
            CHECKING_FOR_CHILDREN = 0x0010,
            FOLDERS = 0x0020,
            NONFOLDERS = 0x0040,
            INCLUDEHIDDEN = 0x0080,
            INIT_ON_FIRST_NEXT = 0x0100,
            NETPRINTERSRCH = 0x0200,
            SHAREABLE = 0x0400,
            STORAGE = 0x0800,
            NAVIGATION_ENUM = 0x1000,
            FASTITEMS = 0x2000,
            FLATLIST = 0x4000,
            ENABLE_ASYNC = 0x8000
        }



        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214F2-0000-0000-C000-000000000046")]
        public interface IEnumIDList
        {
            [PreserveSig]
            HResult Next(uint celt, out ItemIdListSafeHandle rgelt, out uint pceltFetched);

            [PreserveSig]
            HResult Skip(uint celt);

            void Reset();

            IEnumIDList Clone();
        }

        [Flags]
        public enum SFGAO : uint
        {
            CANCOPY = 0x1,
            CANMOVE = 0x2,
            CANLINK = 0x4,
            STORAGE = 0x00000008,
            CANRENAME = 0x00000010,
            CANDELETE = 0x00000020,
            HASPROPSHEET = 0x00000040,
            DROPTARGET = 0x00000100,
            CAPABILITYMASK = 0x00000177,
            ENCRYPTED = 0x00002000,
            ISSLOW = 0x00004000,
            GHOSTED = 0x00008000,
            LINK = 0x00010000,
            SHARE = 0x00020000,
            READONLY = 0x00040000,
            HIDDEN = 0x00080000,
            DISPLAYATTRMASK = 0x000FC000,
            FILESYSANCESTOR = 0x10000000,
            FOLDER = 0x20000000,
            FILESYSTEM = 0x40000000,
            HASSUBFOLDER = 0x80000000,
            CONTENTSMASK = 0x80000000,
            VALIDATE = 0x01000000,
            REMOVABLE = 0x02000000,
            COMPRESSED = 0x04000000,
            BROWSABLE = 0x08000000,
            NONENUMERATED = 0x00100000,
            NEWCONTENT = 0x00200000,
            CANMONIKER = 0x00400000,
            HASSTORAGE = 0x00400000,
            STREAM = 0x00400000,
            STORAGEANCESTOR = 0x00800000,
            STORAGECAPMASK = 0x70C50008,
            PKEYSFGAOMASK = 0x81044000
        }

        [Flags]
        public enum SHGDN
        {
            SHGDN_NORMAL = 0x0000,
            SHGDN_INFOLDER = 0x0001,
            SHGDN_FOREDITING = 0x1000,
            SHGDN_FORADDRESSBAR = 0x4000,
            SHGDN_FORPARSING = 0x8000
        }
    }
}

using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

// ReSharper disable UnassignedReadonlyField
// ReSharper disable InconsistentNaming
#pragma warning disable 169

namespace Example.Native
{
    public static class Gdi32
    {
        public abstract class GdiObjectSafeHandle : SafeHandle
        {
            protected GdiObjectSafeHandle() : base(IntPtr.Zero, true)
            {
            }

            public override bool IsInvalid => handle == IntPtr.Zero;

            protected override bool ReleaseHandle() => DeleteObject(handle);

            [DllImport("gdi32.dll")]
            public static extern bool DeleteObject(IntPtr hObject);
        }

        public sealed class BitmapSafeHandle : GdiObjectSafeHandle
        {
            private BitmapSafeHandle()
            {
            }

            public unsafe Bitmap AsBitmap()
            {
                if (GetObject(this, sizeof(BITMAP), out var info) != sizeof(BITMAP))
                {
                    if (IsInvalid) throw new ObjectDisposedException(nameof(BitmapSafeHandle));
                    throw new Win32Exception("GetObject failed.");
                }

                if (info.bmBits == IntPtr.Zero) throw new InvalidOperationException("Cannot access image data directly without first making a copy.");

                return info.bmHeight > 0
                    ? new Bitmap(info.bmWidth, info.bmHeight, -info.bmWidthBytes, GetPixelFormat(info.bmBitsPixel), info.bmBits + (info.bmHeight - 1) * info.bmWidthBytes)
                    : new Bitmap(info.bmWidth, -info.bmHeight, info.bmWidthBytes, GetPixelFormat(info.bmBitsPixel), info.bmBits);
            }

            public Bitmap CopyAndFree()
            {
                using (this) return new Bitmap(AsBitmap());
            }

            private static PixelFormat GetPixelFormat(ushort numBitsPerPixel)
            {
                switch (numBitsPerPixel)
                {
                    case 32:
                        return PixelFormat.Format32bppArgb;
                    case 24:
                        return PixelFormat.Format24bppRgb;
                    default:
                        throw new NotImplementedException();
                }
            }
        }

        [DllImport("gdi32.dll")]
        public static extern int GetObject(BitmapSafeHandle hBitmap, int cbBuffer, out BITMAP lpBitmap);

        public struct BITMAP
        {
            private readonly int bmType;
            public readonly int bmWidth;
            public readonly int bmHeight;
            public readonly int bmWidthBytes;
            private readonly ushort bmPlanes;
            public readonly ushort bmBitsPixel;
            public readonly IntPtr bmBits;
        }
    }
}

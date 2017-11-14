﻿using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace Example.Native
{
    [DebuggerDisplay("{ToString(),nq}")]
    public struct HResult : IEquatable<HResult>
    {
        public static readonly HResult Ok = (HResult)0;
        public static readonly HResult False = (HResult)1;


        private readonly uint value;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private HResult(uint value)
        {
            this.value = value;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static explicit operator HResult(int value) => new HResult(unchecked((uint)value));

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static explicit operator HResult(uint value) => new HResult(value);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static explicit operator int(HResult hresult) => unchecked((int)hresult.value);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static explicit operator uint(HResult value) => value.value;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static bool operator ==(HResult left, HResult right) => left.value == right.value;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static bool operator !=(HResult left, HResult right) => left.value != right.value;

        public WinErrorCode Code
        {
            get { return unchecked((WinErrorCode)value); }
        }

        public bool IsError
        {
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            get { return (value & 0x80000000) != 0; }
        }

        public HResultFacility Facility
        {
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            get { return unchecked((HResultFacility)((value >> 16) & 0x7FF)); }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public override string ToString()
        {
            return $"0x{value:X8}";
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public Exception CreateException() => new Win32Exception((int)Code);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public Exception CreateException(string message) => new Win32Exception((int)Code, message);

        [DebuggerNonUserCode]
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public void ThrowIfError()
        {
            if (IsError) throw CreateException();
        }

        [DebuggerNonUserCode]
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public void ThrowIfError(string message)
        {
            if (IsError) throw CreateException(message);
        }

        public bool Equals(HResult other) => value == other.value;

        public override bool Equals(object obj) => obj is HResult && Equals((HResult)obj);

        public override int GetHashCode() => (int)value;
    }
}

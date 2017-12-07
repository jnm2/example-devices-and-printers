using System.Diagnostics;
using System.Drawing;

namespace Example.Native
{
    [DebuggerDisplay("{ToString(),nq}")]
    public struct POINT
    {
        public int X;
        public int Y;

        public POINT(int x, int y)
        {
            X = x;
            Y = y;
        }

        public override string ToString()
        {
            return X + ", " + Y;
        }

        public static implicit operator Point(POINT point)
        {
            return new Point(point.X, point.Y);
        }

        public static implicit operator POINT(Point point)
        {
            return new POINT(point.X, point.Y);
        }
    }
}

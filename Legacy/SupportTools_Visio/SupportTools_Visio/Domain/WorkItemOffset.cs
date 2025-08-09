using System.Windows;

namespace SupportTools_Visio.Domain
{
    public class WorkItemOffset
    {
        private int _count;

        private double _rowOffset;

        private double _x;

        private double _xInitial;

        private double _y;

        private double _yInitial;

        public WorkItemOffset(Point initialOffset, double overflowOffset, double padX = 0.05, double padY = 0.05)
        {
            _x = _xInitial = initialOffset.X;
            _y = _yInitial = initialOffset.Y;

            _rowOffset = overflowOffset;

            PadX = padX;
            PadY = padY;
        }

        public int Count
        {
            get => _count;
            set => _count = value;
        }

        public double PadX { get; set; }
        public double PadY { get; set; }

        public double RowOffset
        {
            get => _rowOffset;
            set => _rowOffset = value;
        }

        public double X
        {
            get => _x;
            set => _x = value;
        }

        public double Y
        {
            get => _y;
            set => _y = value;
        }

        public void DecrementHorizontal(double offset)
        {
            //if (Count % 5 == 0)
            //{
            //    _y += RowOffset;
            //    _y += RowOffset > 0 ? PadY : -PadY;
            //    _x = _xInitial;
            //}

            _x -= offset;
            _x -= PadX;

            _count++;
        }

        public void DecrementHorizontal(double offset, OffsetDirection offsetDirection, int columns = 5)
        {
            if (Count % columns == 0)
            {
                switch (offsetDirection)
                {
                    case OffsetDirection.Up:
                        _y += RowOffset + PadY;
                        break;

                    case OffsetDirection.Down:
                        _y -= RowOffset + PadY;
                        break;

                    case OffsetDirection.Left:
                        break;

                    case OffsetDirection.Right:
                        break;
                }

                //if (offsetDirection == OffsetDirection.Up)
                //{
                //    _y += RowOffset + PadY;
                //}
                //else
                //{
                //    _y -= RowOffset + PadY;
                //}

                _x = _xInitial;
            }

            _x -= offset + PadX;

            _count++;
        }

        public void IncrementHorizontal(double offset)
        {
            //if (Count % 5 == 0)
            //{
            //    _y += RowOffset;
            //    _y += RowOffset > 0 ? PadY : -PadY;
            //    _x = _xInitial;
            //}

            _x += offset;
            _x += PadX;

            _count++;
        }

        public void IncrementHorizontal(double offset, OffsetDirection offsetDirection, int columns = 5)
        {
            if (Count % columns == 0)
            {
                switch (offsetDirection)
                {
                    case OffsetDirection.Up:
                        _y += RowOffset + PadY;
                        break;

                    case OffsetDirection.Down:
                        _y -= RowOffset + PadY;
                        break;

                    case OffsetDirection.Left:
                        break;

                    case OffsetDirection.Right:
                        break;
                }

                //if (offsetDirection == OffsetDirection.Up)
                //{
                //    _y += RowOffset + PadY;
                //}
                //else
                //{
                //    _y -= RowOffset + PadY;
                //}

                _x = _xInitial;
            }

            _x += offset + PadX;

            _count++;
        }
    }
}
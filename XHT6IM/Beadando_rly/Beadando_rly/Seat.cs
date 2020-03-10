using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beadando_rly
{
    class Seat
    {
        bool occupied;
        public Seat(bool _occupied)
        {
            occupied = _occupied;
        }
        void setOcuppied(bool b)
        {
            occupied = b;
        }

        bool getOccupied()
        {
            return occupied;
        }
    }
}

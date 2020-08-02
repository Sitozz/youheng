using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace youheng
{
    public class RowaComparer : IEqualityComparer<Rowa>
    {
        public bool Equals(Rowa x, Rowa y)
        {
            return x.Id == y.Id;
        }

        public int GetHashCode(Rowa obj)
        {
            return obj.Id.GetHashCode();
        }
    }
}

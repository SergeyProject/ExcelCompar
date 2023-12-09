using System.Collections.Generic;

namespace ExcelCompar.Models
{

    public class ModelComparer : IEqualityComparer<Model>
    {      
        bool IEqualityComparer<Model>.Equals(Model x, Model y)
        {
            return x.FirstName == y.FirstName && x.SecondName == y.SecondName && x.ThirdName == y.ThirdName;
        }

        int IEqualityComparer<Model>.GetHashCode(Model obj)
        {
           return obj.FirstName.GetHashCode() ^ obj.SecondName.GetHashCode() ^ obj.ThirdName.GetHashCode();
        }
    }

}

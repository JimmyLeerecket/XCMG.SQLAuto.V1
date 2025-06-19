using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XCMG.SQLAuto.V1
{
    public class Model
    {
    }

    public class LookupEntityModel
    {
        public string EntityName { get; set; }

        public string Key { get; set; }

        public string Value { get; set; }

        public string SourceType {  get; set; }
    }

    public class LookupEntityModels
    {
        public List<LookupEntityModel> OldModels { get; set; } = new List<LookupEntityModel>();
        public List<LookupEntityModel> NewModels { get; set; } = new List<LookupEntityModel>();
    }

    public class Product
    {
        public string? Color { get; set; }
        public decimal Price { get; set; }
        public string? Name { get; set; }
        public string? Category { get; set; }
        public string? Size { get; set; }
    }

    public class Category
    {
        public string? Name { get; set; }
        public string? Description { get; set; }
    }
}

using Exceleration.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace Exceleration.Conditions
{
    public class WhereOperator<T> : ICondition
    {
        public string PropertyName { get; set; }
        public T PropertyValue { get; set; }

        public WhereOperator(string propertyName, T propertyValue)
        {
            PropertyName = propertyName;
            PropertyValue = propertyValue;
        }
    }
}

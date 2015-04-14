using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExamReport
{
    class CustomRelation
    {
        public string _Column_name;
        public string _relation;
        public string _value;

        public string Column_type;

        public CustomRelation(string name, string relation, string value)
        {
            _Column_name = name;
            _relation = relation;
            _value = value;
        }

        public string get_string()
        {
            return " " + _Column_name + " " + translate(_relation) + " " + _value + " ";
        }

        public string translate(string relation)
        {
            switch (relation)
            {
                case "等于":
                    return "=";
                case "大于":
                    return ">";
                case "小于":
                    return "<";
                case "大于等于":
                    return ">=";
                case "小于等于":
                    return "<=";
                case "不等于":
                    return "<>";
                default:
                    return "";

            }
        }
    }
}

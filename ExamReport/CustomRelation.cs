using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExamReport
{
    class CustomRelation
    {
        public string _tag;

        public Stack<string> _relation = new Stack<string>();

        public string Column_type;

        public void set_tag(string tag)
        {
            _tag = tag;
        }
        public void insert(string temp)
        {
            _relation.Push(temp);
        }
        public void reset()
        {
            _relation.Clear();
        }
        public bool check(string temp_tag)
        {
            return _tag.Equals(temp_tag);
        }

        public void revoke()
        {
            _relation.Pop();
        }

        public bool isEmpty()
        {
            return _relation.Count == 0;
        }

        public string get_relation()
        {
            StringBuilder sb = new StringBuilder();
            while (_relation.Count != 0)
            {
                sb.Insert(0, _relation.Pop());
                if(_relation.Count != 0)
                    sb.Insert(0, " ");
            }
            return sb.ToString();
        }
    }
}

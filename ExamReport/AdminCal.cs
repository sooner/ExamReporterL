using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ExamReport
{
    class AdminCal
    {
        DataTable _data;
        decimal _fullmark;
        string _name;

        public Admin_WordData w_result;
        public Admin_WordData l_result;

        public AdminCal(Configuration config, DataTable data, decimal fullmark, string name)
        {
            _data = data;
            _fullmark = fullmark;
            _name = name;

            w_result = new Admin_WordData();
            l_result = new Admin_WordData();

        }

        public void Calculate()
        {
            DataTable w_data = _data.equalfilter("type", "w");
            DataTable l_data = _data.equalfilter("type", "l");

            single_process(w_data, w_result);
            single_process(l_data, l_result);
        }

        public void single_process(DataTable data, Admin_WordData result)
        {
            result.total.totalnum = data.Rows.Count;
            result.total.fullmark = _fullmark;
            result.total
        }

    }
}

using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Data;

namespace ExamReport
{
    public class PartitionData
    {
        public DataTable _standard_ans;
        public DataTable _group_ans;
        public string name;
        public int total_num;
        public decimal fullmark;
        public decimal max;
        public decimal min;
        public decimal avg;
        public decimal stDev;
        public decimal Dfactor;
        public decimal difficulty;
        public decimal discriminant;

        public List<decimal> total_discriminant = new List<decimal>();
        public List<decimal> group_discriminant = new List<decimal>();
        public int PLN;
        public int PHN;
        public List<Disc> total = new List<Disc>();
        public List<Disc> group = new List<Disc>();

        
        public DataTable Total_tuple_analysis;

        public DataTable total_analysis;
        public DataTable groups_analysis;
        public DataTable freq_analysis;

        public DataTable total_dist;
        public class group_data
        {
            public DataTable group_dist;
            public DataTable group_detail;
        }

        public ArrayList single_group_analysis;

        public class single_data
        {
            public WordData.single_type stype;
            public DataTable single_detail;
        }

        public ArrayList single_topic_analysis;


        public PartitionData(string _name)
        {
            name = _name;
            total_num = 0;
            fullmark = 0.0m;
            max = 0.0m;
            min = 0.0m;
            avg = 0.0m;
            stDev = 0.0m;
            Dfactor = 0.0m;
            difficulty = 0.0m;
            discriminant = 0.0m;

            total_analysis = new DataTable();
            groups_analysis = new DataTable();
            freq_analysis = new DataTable();
            Total_tuple_analysis = new DataTable();

            single_group_analysis = new ArrayList();
            single_topic_analysis = new ArrayList();
            total_dist = new DataTable();

            total_analysis.Columns.Add("number", typeof(string));
            total_analysis.Columns.Add("total_num", typeof(int));
            total_analysis.Columns.Add("fullmark", typeof(decimal));
            total_analysis.Columns.Add("max", typeof(decimal));
            total_analysis.Columns.Add("min", typeof(decimal));
            total_analysis.Columns.Add("avg", typeof(decimal));
            total_analysis.Columns.Add("stDev", typeof(decimal));
            total_analysis.Columns.Add("dfactor", typeof(decimal));
            total_analysis.Columns.Add("difficulty", typeof(decimal));
            total_analysis.PrimaryKey = new DataColumn[] { total_analysis.Columns["number"] };

            groups_analysis.Columns.Add("number", typeof(string));
            groups_analysis.Columns.Add("fullmark", typeof(decimal));
            groups_analysis.Columns.Add("max", typeof(decimal));
            groups_analysis.Columns.Add("min", typeof(decimal));
            groups_analysis.Columns.Add("avg", typeof(decimal));
            groups_analysis.Columns.Add("stDev", typeof(decimal));
            groups_analysis.Columns.Add("dfactor", typeof(decimal));
            groups_analysis.Columns.Add("difficulty", typeof(decimal));
            groups_analysis.PrimaryKey = new DataColumn[] { groups_analysis.Columns["number"] };

            freq_analysis.Columns.Add("totalmark", typeof(decimal));
            freq_analysis.Columns.Add("frequency", typeof(int));
            freq_analysis.Columns.Add("rate", typeof(decimal));
            freq_analysis.Columns.Add("accumulateFreq", typeof(int));
            freq_analysis.Columns.Add("accumulateRate", typeof(decimal));
            //freq_analysis.PrimaryKey = new DataColumn[] { freq_analysis.Columns["totalmark"] };

            
        }

        public class Disc
        {
            decimal PLN;
            decimal PHN;
            int count;
            decimal fullmark;

            public Disc(int _count, decimal _fullmark)
            {
                PLN = 0m;
                PHN = 0m;
                count = Convert.ToInt32(Math.Ceiling(_count * 0.27));
                fullmark = _fullmark;
            }

            public void AddData(decimal value, bool isPLN)
            {
                if (isPLN)
                    PLN += value;
                else
                    PHN += value;
            }
            public decimal GetAns()
            {
                if (count > 0)
                    return ((PHN - PLN) / count) / fullmark;
                else
                    return 0;
            }
        }
    }
}

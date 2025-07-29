using SQLManage.Models;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQLManage.Util
{
    internal class SqlAccess
    {
        /// <summary>
        /// 系统数据访问
        /// </summary>
        public SqlSugarClient SystemDataAccess;



        public SqlAccess()
        {

            MySqlLink();
        }



        public void MySqlLink()
        {
            //用于存储系统数据的数据库
            SystemDataAccess = new SqlSugarClient(new ConnectionConfig()
            {
                //ConnectionString = "server=127.0.0.1;uid=root;pwd=123456;database=csdk_zj",
                ConnectionString = "server=127.0.0.1;uid=root;pwd=Csdk@2025;database=csdk_zj",
                DbType = DbType.MySql,
                InitKeyType = InitKeyType.Attribute,
                IsAutoCloseConnection = true
            });
            SystemDataAccess.DbMaintenance.CreateDatabase();





        }


        /// <summary>
        /// 初始化表格
        /// </summary>
        public void InitializeTable()
        {
            
            SystemDataAccess.CodeFirst.InitTables(typeof(Tbl_productiondatamodel));
            
        }
    }
}

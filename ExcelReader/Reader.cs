using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;


namespace ExcelReader
{
    class Reader
    {
        public Reader()
        {

        }

        public void ReadPlayers()
        {
            string fileName = @"D:\erdi\players.xlsx";

            string connString = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=Excel 12.0";

            var playerList = new List<Player>();

            OleDbConnection oledbConn = new OleDbConnection(connString);
            try
            {
                // Open connection
                oledbConn.Open();
                // Create OleDbCommand object and select data from worksheet Sheet1
                //string excelSheetName = GetCurrentMonth() + DateTime.Today.Year.ToString() + "SMarka";

                DataTable sheetTable = oledbConn.GetSchema("Tables");
                DataRow rowSheetName = sheetTable.Rows[0];
                string sheetName = sheetTable.Rows[0]["TABLE_NAME"].ToString(); // First Sheet
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [" + sheetName + "]", oledbConn);

                using (OleDbDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        playerList.Add(GetPlayer(dr));
                    }
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                // Close connection
                oledbConn.Close();
            };
        }

        public void ReadTeams()
        {
            string fileName = @"D:\erdi\teams.xlsx";

            string connString = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=Excel 12.0";

            var teamList = new List<Team>();

            OleDbConnection oledbConn = new OleDbConnection(connString);
            try
            {
                // Open connection
                oledbConn.Open();
                // Create OleDbCommand object and select data from worksheet Sheet1
                //string excelSheetName = GetCurrentMonth() + DateTime.Today.Year.ToString() + "SMarka";

                DataTable sheetTable = oledbConn.GetSchema("Tables");
                DataRow rowSheetName = sheetTable.Rows[0];
                string sheetName = sheetTable.Rows[0]["TABLE_NAME"].ToString(); // First Sheet
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [" + sheetName + "]", oledbConn);

                using (OleDbDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        teamList.Add(GetTeam(dr));
                    }
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                // Close connection
                oledbConn.Close();
            };
        }

        private Player GetPlayer(OleDbDataReader dr)
        {
            var player = new Player
            {
                Id = dr["Id"].ToString(),
                Number = dr["Number"].ToString(),
                TeamId = dr["TeamId"].ToString(),
                Name = dr["Name"].ToString()
            };

            return player;
        }

        private Team GetTeam(OleDbDataReader dr)
        {
            var team = new Team
            {
                Id = dr["Id"].ToString(),
                LeaugeId = dr["LeaugeId"].ToString(),
                Name = dr["Name"].ToString()
            };

            return team;
        }
    }
}

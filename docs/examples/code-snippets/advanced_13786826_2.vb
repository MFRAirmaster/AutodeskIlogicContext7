' Title: Could not load System.Data.OleDb (9.0.0.8) in .NET 8 Add-In for Inventor 2025
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/could-not-load-system-data-oledb-9-0-0-8-in-net-8-add-in-for/td-p/13786826
' Category: advanced
' Scraped: 2025-10-07T13:40:03.285882

using System;
using System.Data.OleDb;

namespace MD_PowerTools.Helpers
{
    public class ProvisionalCodeGenerator
    {
        private static readonly char[] Base36Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789".ToCharArray();
        private static readonly char[] Base26Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();

        public int GetUserId(string username) => GetUserIdFromAccess(username);

        public string GenerateCode(int userId)
        {
            string userCode = ToBaseN(userId - 1, Base36Chars, 2);
            DateTime startDate = new DateTime(2025, 1, 1);
            int dayNumber = (DateTime.Now.Date - startDate).Days + 1;
            string dayCode = ToBaseN(dayNumber, Base36Chars, 3);
            int msOfDay = (int)(DateTime.Now - DateTime.Now.Date).TotalMilliseconds;
            string msCode = ToBaseN(msOfDay, Base26Chars, 6);
            return userCode + dayCode + msCode;
        }

        private int GetUserIdFromAccess(string username)
        {
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=M:\NewPath\tgt_tic_tac.accdb;";
            string query = "SELECT id_usuario FROM t_usuario WHERE username = ?";

            using var connection = new OleDbConnection(connectionString);
            using var command = new OleDbCommand(query, connection);
            command.Parameters.AddWithValue("username", username);

            connection.Open();
            object result = command.ExecuteScalar();
            if (result != null && int.TryParse(result.ToString(), out int id))
                return id;
            else
                throw new Exception("User not found in t_usuario.");
        }

        private static string ToBaseN(long value, char[] alphabet, int length)
        {
            int n = alphabet.Length;
            char[] result = new char[length];
            for (int i = length - 1; i >= 0; i--)
            {
                result[i] = alphabet[(int)(value % n)];
                value /= n;
            }
            return new string(result);
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Npgsql;
using System.Data.Common;
using System.Configuration;
namespace NaumenPenalty.DataManage
{
	class ConnectionFactory
	{
		public static DbConnection CreateConnection()
		{
			var connectionString = ConfigurationManager.ConnectionStrings[1];
			switch (connectionString.ProviderName)
			{
				case "Npgsql": return new NpgsqlConnection(connectionString.ConnectionString); break;
				default:throw new ArgumentException("Некорректное имя провайдера данных"); break;
			}
		}
	}

	class CommandFactory
	{
		public static DbCommand CreateCommand(DbConnection connection,string commandText)
		{
			var command=connection.CreateCommand();
			command.CommandText = commandText;
			return command;
		}
	}
}

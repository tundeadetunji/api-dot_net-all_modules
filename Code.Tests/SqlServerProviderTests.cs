using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Xunit;
using iNovation.Code;

namespace Code.Tests {
    public class SqlServerProviderTests {

        [Fact]
        public void GetParameterPrefix_ReturnsAtSymbol()
        {
            var provider = new SqlServerProvider("FakeConnectionString");
            Assert.Equal("@", provider.GetParameterPrefix());
        }

        [Fact]
        public void CreateCommand_ReturnsSqlCommandWithQueryAndConnection()
        {
            var provider = new SqlServerProvider("FakeConnectionString");

            using (var conn = new SqlConnection()) {
                var cmd = provider.CreateCommand("SELECT 1", conn);

                var sqlCmd = Assert.IsType<SqlCommand>(cmd);
                Assert.Equal("SELECT 1", sqlCmd.CommandText);
                Assert.Same(conn, sqlCmd.Connection);
            }
        }

        [Fact]
        public void CreateParameter_ReturnsSqlParameterWithCorrectValues()
        {
            var provider = new SqlServerProvider("FakeConnectionString");

            var param = provider.CreateParameter("@Id", 42);

            var sqlParam = Assert.IsType<SqlParameter>(param);
            Assert.Equal("@Id", sqlParam.ParameterName);
            Assert.Equal(42, sqlParam.Value);
        }

        [Fact]
        public void CloneParameter_ReturnsIdenticalButDistinctSqlParameter()
        {
            var provider = new SqlServerProvider("FakeConnectionString");

            var original = new SqlParameter {
                ParameterName = "@Age",
                Value = 30,
                SqlDbType = System.Data.SqlDbType.Int,
                Size = 4,
                Direction = System.Data.ParameterDirection.Input,
                IsNullable = true,
                Precision = 0,
                Scale = 0,
                SourceColumn = "Age"
            };

            var clone = provider.CloneParameter(original);

            var clonedParam = Assert.IsType<SqlParameter>(clone);

            Assert.Equal(original.ParameterName, clonedParam.ParameterName);
            Assert.Equal(original.Value, clonedParam.Value);
            Assert.Equal(original.SqlDbType, clonedParam.SqlDbType);
            Assert.Equal(original.Size, clonedParam.Size);
            Assert.Equal(original.Direction, clonedParam.Direction);
            Assert.Equal(original.IsNullable, clonedParam.IsNullable);
            Assert.Equal(original.Precision, clonedParam.Precision);
            Assert.Equal(original.Scale, clonedParam.Scale);
            Assert.Equal(original.SourceColumn, clonedParam.SourceColumn);

            Assert.NotSame(original, clonedParam);
        }


    }
}

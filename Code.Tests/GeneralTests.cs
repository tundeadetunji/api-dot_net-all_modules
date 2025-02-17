using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static iNovation.Code.General;
using Xunit;

namespace Code.Tests {
    public class GeneralTests {

        [Theory]
        [InlineData("John", "Doe", "John Doe")]
        public void FullNameFromNames_HappyCase(string FirstName, string LastName, string expected){
            Assert.Equal(expected, FullNameFromNames(FirstName, LastName));
        }
    }
}

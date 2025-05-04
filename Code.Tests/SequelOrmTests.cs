using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using iNovation.Code;

using Moq;

using Xunit;

namespace Code.Tests {

    public class SequelOrmTests {

        private readonly Mock<IOrm> _ormMock;
        private readonly SequelOrmService _sut;

        public SequelOrmTests()
        {
            _ormMock = new Mock<IOrm>();
            _sut = new SequelOrmService(_ormMock.Object);
        }

        [Fact]
        public void Create_ShouldInsertParentAndChildEntitiesCorrectly()
        {
            // Arrange
            var people = new List<Person>
            {
                new Person
                {
                    Name = "Alice",
                    Phones = new List<Phone>
                    {
                        new Phone { Number = "123" },
                        new Phone { Number = "456" }
                    }
                }
            };

            _ormMock.Setup(m => m.Create(people)).Returns(people);

            // Act
            var result = _sut.Create(people);

            // Assert
            Assert.NotNull(result);
            Assert.Single(result);
            Assert.Equal("Alice", result[0].Name);
            Assert.Equal(2, result[0].Phones.Count);
            _ormMock.Verify(m => m.Create(people), Times.Once);
        }

        [Fact]
        public void Create_ShouldHandleEmptyListGracefully()
        {
            // Arrange
            var people = new List<Person>();

            _ormMock.Setup(m => m.Create(people)).Returns(people);

            // Act
            var result = _sut.Create(people);

            // Assert
            Assert.NotNull(result);
            Assert.Empty(result);
            _ormMock.Verify(m => m.Create(people), Times.Once);
        }

        [Fact]
        public void Create_ShouldThrowExceptionWhenOrmFails()
        {
            // Arrange
            var people = new List<Person> { new Person { Name = "Bob" } };
            _ormMock.Setup(m => m.Create(people)).Throws(new InvalidOperationException("DB failure"));

            // Act & Assert
            var ex = Assert.Throws<InvalidOperationException>(() => _sut.Create(people));
            Assert.Equal("DB failure", ex.Message);
        }

        [Fact]
        public void Create_ShouldThrowArgumentNullExceptionOnNullInput()
        {
            // Act & Assert
            var ex = Assert.Throws<ArgumentNullException>(() => _sut.Create(null));
            Assert.Equal("people", ex.ParamName);
        }
    }

    // Minimal service wrapper to test (replace with your real service if needed)
    public class SequelOrmService {
        private readonly IOrm _orm;

        public SequelOrmService(IOrm orm)
        {
            _orm = orm ?? throw new ArgumentNullException(nameof(orm));
        }

        public List<Person> Create(List<Person> people)
        {
            if (people == null)
                throw new ArgumentNullException(nameof(people));

            return _orm.Create(people);
        }
    }

    public interface IOrm {
        List<Person> Create(List<Person> people);
    }

    public class Person {
        public string Name { get; set; }
        public List<Phone> Phones { get; set; } = new List<Phone>();
    }

    public class Phone {
        public string Number { get; set; }
    }
}

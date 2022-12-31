namespace WebAPI
{
    public class TestExcel
    {
        public string? Name { get; set; }
        public string? Occupation { get; set; }

        public TestExcel()
        {

        }
        public override string ToString()
        {
            return $"{Name} {Occupation}";
        }
    }
}

namespace Excel.TestDoubles
{
    public class Excel2013Facade
    {
        public Excel2013Facade()
        {
            Application = new ApplicationTestDouble();
        }

        public ApplicationTestDouble Application { get; private set; }
    }
}
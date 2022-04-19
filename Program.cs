using Sap;

namespace CSap;
public class Program
{
    public static void Main(string[] args)
    {
        // Testando a classe SapConnection
        SapConnection sapConnection = new SapConnection();
        sapConnection.Connection();
        sapConnection.GetAllSessions();
        Console.WriteLine(sapConnection.SapSessions?.Count);
        sapConnection.Close();
    }
}


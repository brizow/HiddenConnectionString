using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;
using ADODB;

http://stackoverflow.com/questions/3217014/how-to-securely-store-connection-string-details-in-vba
//Since connection strings cannot be hidden in VBA, this little module allows you to connect a VBA
// macro set to a datastore using this hidden ADODB connection string. Call with you variable name.GetConnection()

namespace HiddenConnectionString
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual),
    Guid("2FCEF713-CD2E-4ACB-A9CE-E57E7F51E72E")]
    public interface IMyServer
    {
        Connection GetConnection();
        void Shutdown();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("57BBEC44-C6E6-4E14-989A-B6DB7CF6FBEB")]
    public class MyServer : IMyServer
    {
        private Connection cn;
        //Enter your connection string provider, data source, initial catalog, user, pass
        private string cnStr = "Provider=SQLOLEDB; Data Source=SERVER\\INSTANCE; Initial Catalog=default_catalog; User ID=your_username; Password=your_password";

        public MyServer()
        {
        }

        public Connection GetConnection()
        {
            cn = new Connection();
            cn.ConnectionString = cnStr;
            cn.Open();
            return cn;
        }

        public void Shutdown()
        {
            cn.Close();
        }
    }
}
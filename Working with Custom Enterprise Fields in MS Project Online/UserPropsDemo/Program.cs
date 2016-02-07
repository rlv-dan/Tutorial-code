using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UserPropsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            var myUsername = "myuser@mytenant.onmicrosoft.com";
            var myPassword = "12345";
            var myProjectSite = "https://mytenant.sharepoint.com/sites/PROJ";

            // First create an instance of the class. It will connect to Project Online automatically
            var userProps = new UserProps(myProjectSite, myUsername, myPassword);

            // Now we can call one of the methods:
            userProps.ListContextData();
            userProps.ListUserResourceCustomFields("someuser@mytenant.onmicrosoft.com");
            userProps.SetCustomField("someuser@mytenant.onmicrosoft.com", "New Value");

            Console.WriteLine("\n\nReady!");
            Console.ReadKey();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.DirectoryServices;
using System.Activities;
using System.ComponentModel;
using System.Net.NetworkInformation;

namespace ActiveDirectory
{
    public class ActiveDirectoryclass : CodeActivity
    {
        
        public string dominName = Environment.UserDomainName;
        public string userName = Environment.UserName;
        public string strError = string.Empty;               
        public DataTable datatablelist = new DataTable();
        
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> ADPath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> Email { get; set; }

        [Category("Output")]
        public OutArgument<DataTable> Userinfolist { get; set; }
        
        protected override void Execute(CodeActivityContext context)
        {
            var adPath = ADPath.Get(context);
            var email = Email.Get(context);

            strError = "";
            string domainAndUsername = dominName + @"\" + userName;
            DirectoryEntry entry = new DirectoryEntry(adPath, domainAndUsername, "", AuthenticationTypes.ReadonlyServer);

            // This varifies the user with Active Directory and if user is not valid then exception is thrown.
            object obj = entry.NativeObject;

            if (obj != null)
            {               
                DirectorySearcher search = new DirectorySearcher();               

                search.Filter = string.Format("(&(objectClass=user)(objectCategory=user) (mail={0}))", email);               
                search.SearchRoot = entry;
                search.SearchScope = SearchScope.Subtree;                
                search.Sort = new SortOption("givenName", SortDirection.Ascending);
                SearchResultCollection results = search.FindAll();

                foreach (SearchResult r in results)
                {
                    ResultPropertyCollection props = r.Properties;
                                        
                    datatablelist.Columns.Add("PropertName");
                    datatablelist.Columns.Add("PropertyValue");


                    foreach (string propName in props.PropertyNames)
                    {                        
                        DataRow dr = datatablelist.NewRow();
                        dr["PropertName"] = propName;
                        dr["PropertyValue"] = props[propName][0];
                        datatablelist.Rows.Add(dr);
                    }

                }                
                Userinfolist.Set(context, datatablelist);               
            }
        }
    }
}



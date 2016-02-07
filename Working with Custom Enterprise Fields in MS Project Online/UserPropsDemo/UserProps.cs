using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System.Security;

namespace UserPropsDemo
{
    class UserProps
    {
        private ProjectContext context;

        // Custom Enterprise Fields
        private const string CustomFieldGuid = "618D383A-CCBD-11E5-9FC3-966EB9876117";      // must be a valid guid for a custom field in your MS Project Online
        private Dictionary<string, CustomFieldEntity> AllCustomFields = null;

        // Internal classes for keeping data
        private class LookupEntryEntity
        {
            public string Id { get; set; }
            public string InternalName { get; set; }
            public string Value { get; set; }
        }

        private class CustomFieldEntity
        {
            public string Id { get; set; }
            public string InternalName { get; set; }
            public string Name { get; set; }

            public Dictionary<string, LookupEntryEntity> LookupEntries { get; set; }
            public CustomFieldEntity()
            {
                LookupEntries = new Dictionary<string, LookupEntryEntity>();
            }
        }

        // Constructor. Will connect to Project Online with mandratory credentials.
        public UserProps(string site, string username, string password)
        {
            try
            {
                if (site == "" || username == "" || password == "")
                {
                    throw (new Exception("Must supply site, username and password!"));
                }

                Console.WriteLine("Connecting to Project Online @ " + site + "...");

                var securePassword = new SecureString();
                foreach (var ch in password.ToCharArray())
                {
                    securePassword.AppendChar(ch);
                }

                context = new ProjectContext(site);
                context.Credentials = new SharePointOnlineCredentials(username, securePassword);
                context.Load(context.Web);
                context.ExecuteQuery();
                Console.WriteLine("   Connected to '" + context.Web.Title + "'");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        // Public facing function to set the custom enterprise resource fields
        public void SetCustomField(string userEmail, string newValue = "")
        {
            var user = GetUserResource(userEmail);
            if (user != null)
            {
                // Get all custom fields
                LoadCustomFields();

                // Set the custom enterprise resource field
                try
                {
                    Console.WriteLine("Setting custom enterprise resource property for user...");
                    newValue = newValue.ToLower();
                    var fieldInternalName = AllCustomFields[CustomFieldGuid].InternalName;
                    var entryInternalName = AllCustomFields[CustomFieldGuid].LookupEntries[newValue].InternalName;  // not that we are doing a string match here to figure out the internal name of the value in the lookup table
                    user[fieldInternalName] = new string[] { entryInternalName };   // note that it should be a string array
                    Console.WriteLine("\t" + AllCustomFields[CustomFieldGuid].Name + " >> " + AllCustomFields[CustomFieldGuid].LookupEntries[newValue].Value);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Could not set custom field " + CustomFieldGuid + " to " + newValue + ": " + ex.Message);
                }


                // Persist changes (note than the resource object is in the EnterpriseResources collection)
                try
                {
                    Console.WriteLine("\tSaving changes...");
                    context.EnterpriseResources.Update();
                    context.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error saving: " + ex.Message);
                }
            }
        }

        // Load all custom enterprise resource fields and package in an easy to access dictionary.
        private void LoadCustomFields()
        {
            if (AllCustomFields == null)
            {
                Console.WriteLine("Loading custom fields and lookup entries...");

                // In this example I am only using one custom field. You can add more custom field guids to this list to handle multiple fields
                var customFields = new List<CustomField> { context.CustomFields.GetByGuid(new Guid(CustomFieldGuid)) };
                foreach (var field in customFields)
                {
                    context.Load(field);
                    context.Load(field.LookupEntries);
                }
                context.ExecuteQuery();

                // Package custom fields in an easy to access format
                AllCustomFields = new Dictionary<string, CustomFieldEntity>();
                foreach (var field in customFields)
                {
                    //Console.WriteLine(field.InternalName + " = " + field.Name);
                    var cfe = new CustomFieldEntity() { Id = field.Id.ToString(), InternalName = field.InternalName, Name = field.Name };
                    foreach (var entry in field.LookupEntries)
                    {
                        //Console.WriteLine("\t" + entry.InternalName + " = " + entry.FullValue);
                        cfe.LookupEntries.Add(
                                                entry.FullValue.ToLower(),
                                                new LookupEntryEntity() { Id = entry.Id.ToString(), InternalName = entry.InternalName, Value = entry.FullValue }
                                             );
                    }
                    AllCustomFields.Add(field.Id.ToString(), cfe);
                }
            }
        }


        // Loads a user as an enterprise resouce
        private EnterpriseResource GetUserResource(string userEmail)
        {
            try
            {
                Console.WriteLine("Loading user resource for '" + userEmail + "'");

                // Since we can't trust that email is synced to project, get user by login name instead
                string claimsPrefix = "i:0#.f|membership|";
                var loginName = claimsPrefix + userEmail;
                User user = context.Web.SiteUsers.GetByLoginName(loginName);
                EnterpriseResource res = context.EnterpriseResources.GetByUser(user);
                context.Load(res);
                context.ExecuteQuery();
                Console.WriteLine("   Got resource: " + res.Name + " {" + res.Id + "}");
                return res;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error loading user: " + ex.Message);
                return null;
            }
        }


        // --------------- Below are some other examples on how to get data from Project ---------------


        // This shows how to get all custom fields present on a user
        public void ListUserResourceCustomFields(string email)
        {
            var userRes = GetUserResource(email);

            // iterate custom fields
            Console.WriteLine("Loading custom fields...");

            var allFields = new Dictionary<string, Dictionary<string, string>>();

            var customFields = userRes.CustomFields;
            context.Load(customFields);
            context.ExecuteQuery();

            foreach (var field in customFields)
            {
                context.Load(field);
                context.Load(field.LookupEntries);
                context.ExecuteQuery();

                //Console.WriteLine("\t" +  field.Name + " {" + field.InternalName + "}");
                var entries = new Dictionary<string, string>();
                entries.Add("KEYNAME", field.Name);

                foreach (var entry in field.LookupEntries)
                {
                    //Console.WriteLine("\t  " + entry.FullValue + " {" + entry.InternalName + "}");
                    entries.Add(entry.InternalName, entry.FullValue);
                }

                allFields.Add(field.InternalName, entries);
            }

            Console.WriteLine("-----------------User Custom Fields-----------------");
            var fieldValues = userRes.FieldValues;
            if (fieldValues.Count == 0)
            {
                Console.WriteLine("User has no custom fields...");
            }
            foreach (var fieldValue in fieldValues)
            {
                Console.WriteLine(allFields[fieldValue.Key]["KEYNAME"] + " {" + fieldValue.Key + "}");

                foreach (var value in (string[])fieldValue.Value)
                {
                    Console.WriteLine("\t" + allFields[fieldValue.Key][value] + " {" + value + "}");
                }
            }
        }

        // This function shows how to list facts about the MS Project site
        public void ListContextData()
        {
            Console.WriteLine("\nListing context data...\n");

            Console.WriteLine("---------------All Projects---------------");
            context.Load(context.Projects);
            context.ExecuteQuery();

            foreach (PublishedProject proj in context.Projects)
            {
                Console.WriteLine(proj.Name);
            }

            Console.WriteLine("---------------All Site Users---------------");
            UserCollection siteUsers = context.Web.SiteUsers;
            context.Load(siteUsers);
            context.ExecuteQuery();

            var peopleManager = new PeopleManager(context);
            foreach (var user in siteUsers)
            {
                try
                {
                    PersonProperties userProfile = peopleManager.GetPropertiesFor(user.LoginName);
                    context.Load(userProfile);
                    context.ExecuteQuery();

                    Console.WriteLine(userProfile.DisplayName + " (" + userProfile.Email + ")");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error loading user: " + user.LoginName + " --> " + ex.Message);
                }
            }


            Console.WriteLine("---------------All Site User Resources---------------");
            EnterpriseResourceCollection resources = context.EnterpriseResources;
            context.Load(resources);
            context.ExecuteQuery();

            foreach (EnterpriseResource res in resources)
            {
                Console.WriteLine(res.Name + " {" + res.Id + "}");
            }

            Console.WriteLine("---------------Custom Fields---------------");
            CustomFieldCollection customFields = context.CustomFields;
            context.Load(customFields);
            context.ExecuteQuery();
            foreach (CustomField cf in customFields)
            {
                Console.WriteLine(cf.Name + " {" + cf.Id + "}");
            }

            Console.WriteLine("---------------Lookup Tables---------------");
            LookupTableCollection lookupTables = context.LookupTables;
            context.Load(lookupTables);
            context.ExecuteQuery();
            foreach (LookupTable lt in lookupTables)
            {
                Console.WriteLine(lt.Name + " {" + lt.Id + "}");

                context.Load(lt.Entries);
                context.ExecuteQuery();
                foreach (LookupEntry entry in lt.Entries)
                {
                    Console.WriteLine("    " + entry.FullValue + " {" + entry.Id + "}");
                }
            }
        }

    }
}

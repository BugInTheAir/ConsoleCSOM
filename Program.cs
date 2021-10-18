using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Collections.Generic;

namespace ConsoleCSOM
{
    class Program
    {
        private static ISharePointServices _services;
        static async Task Main(string[] args)
        {
            using (var clientContextHelper = new ClientContextHelper())
            {
                ClientContext ctx = GetContext(clientContextHelper);
                _services = new SharePointServices(ctx);
                var existedList = ctx.Web.Lists.GetByTitle("CSOM Test");
                //ctx.Load(ctx.Web, w => w.Fields);
                //await ctx.ExecuteQueryAsync();
                //var myField = ctx.Web.Fields.Where(x => x.Title.Equals("city")).FirstOrDefault();
                //myField.Required = false;
                //myField.Update();
                await ctx.ExecuteQueryAsync();
                for(int i = 0; i < 4; i ++)
                {
                    ListItemCreationInformation oListItemCreationInformation = new ListItemCreationInformation();
                    ListItem oItem = existedList.AddItem(oListItemCreationInformation);
                    var formValues = new List<ListItemFormUpdateValue>();
                    formValues.Add(new ListItemFormUpdateValue() { FieldName = "Title", FieldValue = $"test {i}" });
                    formValues.Add(new ListItemFormUpdateValue() { FieldName = "aboutCT", FieldValue = $"about test {i}" });
                    formValues.Add(new ListItemFormUpdateValue() { FieldName = "cityInfo", FieldValue = $"city test {i}" });
                    oItem.ValidateUpdateListItem(formValues, true, string.Empty, true, true);
                    oItem.Update();
                }
                existedList.Update();
                await ctx.ExecuteQueryAsync();
                //CreateCSOMTestList(ctx);
                //TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);
                //TermStore termStore = session.GetDefaultSiteCollectionTermStore();
                //TermGroup termGroup = termStore.Groups.GetByName("People");

                //ctx.Load(termGroup, tg => tg.TermSets);
                //ctx.ExecuteQuery();
                //await _services.SaveContextAsync();

                //ctx.Load(ctx.Web);
                //await ctx.ExecuteQueryAsync();

                //await SimpleCamlQueryAsync(ctx);
                //await CsomTermSetAsync(ctx);
            }

            Console.WriteLine($"Press Any Key To Stop!");
            Console.ReadKey();
        }

        private static async Task SetMyContentTypeAsDefault(ClientContext ctx)
        {
            var existedList = ctx.Web.Lists.GetByTitle("CSOM Test");
            var currentOrder = existedList.ContentTypes;
            ctx.Load(currentOrder);
            await ctx.ExecuteQueryAsync();
            IList<ContentTypeId> reverseOrder = new List<ContentTypeId>();
            foreach (var type in currentOrder)
            {
                if (type.Name.Equals("CSOM Test content type"))
                {
                    reverseOrder.Add(type.Id);
                }
            }

            existedList.RootFolder.UniqueContentTypeOrder = reverseOrder;
            existedList.RootFolder.Update();
            existedList.Update();
        }

        private static async Task AddContentTypeToMyList(ClientContext ctx)
        {
            ctx.Load(ctx.Web, w => w.Fields, w => w.ContentTypes);
            var existedList = ctx.Web.Lists.GetByTitle("CSOM Test");
            var contentCollection = ctx.Web.ContentTypes;
            await ctx.ExecuteQueryAsync();
            var myContentType = contentCollection.Where(x => x.Name.Equals("CSOM Test content type")).FirstOrDefault();
            FieldLinkCreationInformation info = new FieldLinkCreationInformation();
            info.Field = ctx.Web.Fields.Where(x => x.Title.Equals("about")).FirstOrDefault();
            myContentType.FieldLinks.Add(info);
            info = new FieldLinkCreationInformation();
            info.Field = ctx.Web.Fields.Where(x => x.Title.Equals("city")).FirstOrDefault();
            myContentType.FieldLinks.Add(info);
            existedList.ContentTypes.AddExistingContentType(myContentType);
        }

        private static async Task CreateCustomContentType(ClientContext ctx)
        {
            var contentCollection = ctx.Web.ContentTypes;
            ctx.Load(contentCollection);
            await ctx.ExecuteQueryAsync();
            var parentType = contentCollection.Where(x => x.Name.Equals("Item")).FirstOrDefault();
            ContentTypeCreationInformation oContentTypeCreationInformation = new ContentTypeCreationInformation();

            // Name of the new content type
            oContentTypeCreationInformation.Name = "CSOM Test content type";

            // Description of the new content type
            oContentTypeCreationInformation.Description = "My custom content type created by csom";

            // Name of the group under which the new content type will be creted
            oContentTypeCreationInformation.Group = "Custom Content Types Group";

            // Specify the parent content type over here
            oContentTypeCreationInformation.ParentContentType = parentType;

            // Add "ContentTypeCreationInformation" object created above
            ContentType oContentType = contentCollection.Add(oContentTypeCreationInformation);
            
        }

        private static void CreateCityField(ClientContext ctx)
        {
            ctx.Web.Fields.AddFieldAsXml($@"
<Field
    Type=""TaxonomyFieldType""
    DisplayName = ""city""
    Description = ""city""
    Required = ""FALSE""
    EnforceUniqueValues = ""FALSE""
    Indexed = ""FALSE""
    MaxLength = ""255""
    ID = ""{Guid.NewGuid().ToString()}""
    Name = ""cityInfo"" >
</Field>

                ", true, AddFieldOptions.DefaultValue);
        }
        private static void CreateAboutField(ClientContext ctx)
        {
            ctx.Web.Fields.AddFieldAsXml($@"
<Field
    Type=""Text""
    DisplayName = ""about""
    Description = ""about city""
    Required = ""TRUE""
    EnforceUniqueValues = ""FALSE""
    Indexed = ""FALSE""
    MaxLength = ""255""
    ID = ""{Guid.NewGuid().ToString()}""
    Name = ""aboutCT"" >
</Field>

                ", true, AddFieldOptions.DefaultValue);
        }

        private static void CreateMyTerms()
        {
            _services.CreateTerm("Ho Chi Minh", _services.GetTermSetByName("HCM-LeMinhMan", _services.GetTermGroupByName("CSOM Term")));
            _services.CreateTerm("Stockholm", _services.GetTermSetByName("HCM-LeMinhMan", _services.GetTermGroupByName("CSOM Term")));
        }

        private static void CreateMyTermSet()
        {
            _services.CreateTermSet(_services.CreateTermGroup("CSOM Term"), "HCM-LeMinhMan");
        }

        private static void CreateCSOMTestList(ClientContext ctx)
        {
            _services.CreateList("CSOM Test");
        }

        static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }

        private static async Task GetFieldTermValue(ClientContext Ctx, string termId)
        {
            //load term by id
            TaxonomySession session = TaxonomySession.GetTaxonomySession(Ctx);
            Term taxonomyTerm = session.GetTerm(new Guid(termId));
            //taxonomyTerm.CreateTerm();
            Ctx.Load(taxonomyTerm, t => t.Labels,
                                   t => t.Name,
                                   t => t.Id);
            await Ctx.ExecuteQueryAsync();
        }

        private static async Task ExampleSetTaxonomyFieldValue(ListItem item, ClientContext ctx)
        {
            var field = ctx.Web.Fields.GetByTitle("fieldname");

            ctx.Load(field);
            await ctx.ExecuteQueryAsync();

            var taxField = ctx.CastTo<TaxonomyField>(field);

            taxField.SetFieldValueByValue(item, new TaxonomyFieldValue()
            {
                WssId = -1, // alway let it -1
                Label = "correct label here",
                TermGuid = "term id"
            });
            item.Update();
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CsomTermSetAsync(ClientContext ctx)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get the term store by name
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName("Test");
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName("Test Term Set");

            var terms = termSet.GetAllTerms();

            ctx.Load(terms);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CsomLinqAsync(ClientContext ctx)
        {
            var fieldsQuery = from f in ctx.Web.Fields
                              where f.InternalName == "Test" ||
                                    f.TypeAsString == "TaxonomyFieldTypeMulti" ||
                                    f.TypeAsString == "TaxonomyFieldType"
                              select f;

            var fields = ctx.LoadQuery(fieldsQuery);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task SimpleCamlQueryAsync(ClientContext ctx)
        {
            var list = ctx.Web.Lists.GetByTitle("Documents");

            var allItemsQuery = CamlQuery.CreateAllItemsQuery();
            var allFoldersQuery = CamlQuery.CreateAllFoldersQuery();

            var items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View>
                                <Query>
                                    <OrderBy><FieldRef Name='ID' Ascending='False'/></OrderBy>
                                </Query>
                                <RowLimit>20</RowLimit>
                            </View>",
                FolderServerRelativeUrl = "/sites/test-site-duc-11111/Shared%20Documents/2"
                //example for site: https://omniapreprod.sharepoint.com/sites/test-site-duc-11111/
            });

            ctx.Load(items);
            await ctx.ExecuteQueryAsync();
        }
    }
}

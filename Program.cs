using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Collections.Generic;
using System.Globalization;

namespace ConsoleCSOM
{
    class Program
    {
        private static ISharePointServices _services;
        static async Task Main(string[] args)
        {
            using (var clientContextHelper = new ClientContextHelper())
            {
                ContentTypeHelper contentTypeHelper = new ContentTypeHelper();
                ClientContext ctx = GetContext(clientContextHelper);
                _services = new SharePointServices(ctx);
                var existedList = GetMyList(ctx);
                //SetDefaultValueToTaxonomyCities(ctx);
                //var myTerm = _services.GetTermSetByName(Constants.MY_TERM_SET_NAME, _services.GetTermGroupByName(Constants.MY_TERM_GROUP));
                //ctx.Load(myTerm, term => term.Terms);
               
                
                

                
                
                await _services.SaveContextAsync();

                //ctx.Load(ctx.Web, ctx => ctx.Fields);
                //await AddFieldSpecificListContentType(existedList, ctx.Web.Fields, ctx, new string[] { Constants.AUTHOR_TITLE }, Constants.CSOM_TEST_CONTENT_TYPE, false);
            }

            Console.WriteLine($"Press Any Key To Stop!");
            Console.ReadKey();
        }

        private static async Task UpdateAboutDefaultValue(ClientContext ctx)
        {
            var existedList = ctx.Web.Lists.GetByTitle(Constants.CSOM_TEST);
            var items = await MyAboutViewCAML(ctx);
            foreach (var item in items)
            {
                if ((string)item["aboutCT"] == "about default")
                {
                    item["aboutCT"] = "Update script";
                    item.Update();
                }
            }
        }
        private static void SetDefaultValueToTaxonomyCities(ClientContext context)
        {
            var field = context.Web.Fields.GetByInternalNameOrTitle(Constants.CITIES_FIELD_NAME);
            TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(field);
            context.Load(taxonomyField, t => t.TermSetId);
            context.ExecuteQuery();
            taxonomyField.TermSetId = new Guid("da9af3b7-d98f-4638-bbfe-cb104bd31337");
            taxonomyField.UpdateAndPushChanges(true);
            context.ExecuteQuery();
        }
        private static async Task SetDefaultValueToTaxonomyTerm(ClientContext context)
        {
            string termLabel = "Default term";
            Guid termId = new Guid("{a8cb8104-8c93-4cbd-8486-bd4d902673b3}");
            var field = context.Web.Fields.GetByInternalNameOrTitle(Constants.CITY);

            TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(field);
            context.Load(taxonomyField, t => t.DefaultValue);
            context.ExecuteQuery(); // Get the Taxonomy Field

            TaxonomyFieldValue defaultValue = new TaxonomyFieldValue();
            defaultValue.WssId = -1;
            defaultValue.Label = termLabel;
            // GUID should be stored lowercase, otherwise it will not work in Office 2010
            defaultValue.TermGuid = termId.ToString();

            // Get the Validated String for the taxonomy value
            var validatedValue = taxonomyField.GetValidatedString(defaultValue);
            await context.ExecuteQueryAsync();

            // Set the selected default value for the site column
            taxonomyField.DefaultValue = validatedValue.Value;
            taxonomyField.UserCreated = false;
            taxonomyField.UpdateAndPushChanges(true);
            await context.ExecuteQueryAsync();
        }
        private static async Task CreateMyOrderCreatedByDateView(ClientContext ctx)
        {
            var list = ctx.Web.Lists.GetByTitle(Constants.CSOM_TEST);
            var myQuery = @"<View>
                                <Query>
                                    <OrderBy><FieldRef Name='Created' Ascending='FALSE'/></OrderBy>
                                    <Where>
                                      <Contains>
                                          <FieldRef Name='cityInfo' />
                                          <Value Type='Text'>Ho Chi Minh</Value>
                                      </Contains>
                                    </Where>
                                </Query>
                            </View>";
            string[] myViewFields = {Constants.CITY,Constants.ABOUT, Constants.TITLE};
            ViewCreationInformation creationInformation = new ViewCreationInformation()
            {
                Title = Constants.MY_VIEW_TITLE,
                ViewTypeKind = ViewType.Grid,
                Query = myQuery,
                ViewFields = myViewFields
            };
            list.Views.Add(creationInformation);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task<ListItemCollection> MyAboutViewCAML(ClientContext ctx)
        {
            var list = ctx.Web.Lists.GetByTitle(Constants.CSOM_TEST);
            var items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View>
                                <Query>
                                    <Where>
                                         <Eq>
                                            <FieldRef Name=""aboutCT""></FieldRef>
                                            <Value Type=""Text"">about default</Value>
                                          </Eq>
                                  </Where>
                                </Query>
                            </View>"
            });
            ctx.Load(items);
            await ctx.ExecuteQueryAsync();
            return items;
        }

        private static async Task SetDefaultValueToMyList(ClientContext ctx)
        {
            var existedList = ctx.Web.Lists.GetByTitle(Constants.CSOM_TEST);
            ctx.Load(existedList.ContentTypes);
            await ctx.ExecuteQueryAsync();
            var myContentType = existedList.ContentTypes.Where(x => x.Name.Equals(Constants.CSOM_TEST_CONTENT_TYPE)).FirstOrDefault();
            ctx.Load(myContentType.Fields);
            await ctx.ExecuteQueryAsync();
            var about = myContentType.Fields.Where(x => x.Title.Equals("about")).FirstOrDefault();
            about.DefaultValue = "about default";
            about.Update();
            myContentType.Update(false);
            existedList.Update();
        }

        private static async Task CreateMyItems(ClientContext ctx)
        {
            var existedList = ctx.Web.Lists.GetByTitle(Constants.CSOM_TEST);
            await ctx.ExecuteQueryAsync();
            for (int i = 0; i < 2; i++)
            {
                ListItemCreationInformation oListItemCreationInformation = new ListItemCreationInformation();
                ListItem oItem = existedList.AddItem(oListItemCreationInformation);
                var formValues = new List<ListItemFormUpdateValue>();
                formValues.Add(new ListItemFormUpdateValue() { FieldName = "Title", FieldValue = $"test default {i}" });
                //formValues.Add(new ListItemFormUpdateValue() { FieldName = "cityInfo", FieldValue = $"city test {i}" });
                oItem.ValidateUpdateListItem(formValues, true, string.Empty, true, true);
                oItem.Update();
            }
            existedList.Update();
        }

     

        private static List GetMyList(ClientContext ctx)
        {
            return ctx.Web.Lists.GetByTitle(Constants.CSOM_TEST);
        }

     

        private static void CreateCitiesField(ClientContext ctx)
        {
            ctx.Web.Fields.AddFieldAsXml($@"
<Field
    Type=""TaxonomyFieldTypeMulti""
    DisplayName = ""cities""
    Description = ""cities""
    Mult = ""TRUE""
    Required = ""FALSE""
    EnforceUniqueValues = ""FALSE""
    Indexed = ""FALSE""
    MaxLength = ""255""
    ID = ""{Guid.NewGuid().ToString()}""
    Name = ""{Constants.CITIES_FIELD_NAME}"" >
</Field>

                ", true, AddFieldOptions.DefaultValue);
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
        private static async Task CreateAuthorField(ClientContext ctx)
        {
            ctx.Web.Fields.AddFieldAsXml($@"<Field
  Type=""User""
  DisplayName = ""Author_""
  List = ""UserInfo""
  Required = ""FALSE""
  EnforceUniqueValues = ""FALSE""
  ShowField = ""ImnName""
  UserSelectionMode = ""PeopleOnly""
  UserSelectionScope = ""0""
  ID = ""{Guid.NewGuid().ToString()}""
  StaticName = ""Author""
  ShowInEditForm = ""TRUE""
  ShowInNewForm = ""TRUE""
  Name = ""{Constants.AUTHOR_FIELD_NAME}"" >
</Field> ", true, AddFieldOptions.DefaultValue);
            await ctx.ExecuteQueryAsync();
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
            _services.CreateList(Constants.CSOM_TEST);
        }

        static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }
        private static TermSetCollection GetTermSetCollection(ClientContext ctx, string termTitle) {
            TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);
            return session.GetTermSetsByName(termTitle, CultureInfo.CurrentCulture.LCID);
        }
        private static async Task<Term> GetFieldTermValue(ClientContext Ctx, string termId)
        {
            //load term by id
            TaxonomySession session = TaxonomySession.GetTaxonomySession(Ctx);
            Term taxonomyTerm = session.GetTerm(new Guid(termId));
            //taxonomyTerm.CreateTerm();
            Ctx.Load(taxonomyTerm, t => t.Labels,
                                   t => t.Name,
                                   t => t.Id);
            await Ctx.ExecuteQueryAsync();
            return taxonomyTerm;
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

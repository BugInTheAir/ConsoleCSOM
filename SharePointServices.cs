 using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Globalization;
using System.Threading.Tasks;

namespace ConsoleCSOM
{
    public class SharePointServices : ISharePointServices
    {
        private readonly ClientContext _context;

        public SharePointServices(ClientContext context)
        {
            _context = context;
        }


        public void CreateFolderInDocument(List documents, string folderInternalName, string folderName, string folderUrl)
        {
            ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
            listItemCreationInformation.UnderlyingObjectType = FileSystemObjectType.Folder;
            listItemCreationInformation.LeafName = folderInternalName;
            listItemCreationInformation.FolderUrl = folderUrl;
            ListItem listItem = documents.AddItem(listItemCreationInformation);
            listItem["Title"] = folderName;

            listItem.Update();
        }

        public void CreateDocumentList(string name)
        {
            _context.Web.Lists.Add(new ListCreationInformation()
            {
                Title = name,
                TemplateType = (int)ListTemplateType.DocumentLibrary
            });
        }

        public void CreateList(string name)
        {
            _context.Web.Lists.Add(new ListCreationInformation()
            {
                Title = name,
                TemplateType = (int)ListTemplateType.CustomGrid
            });
        }

        public void CreateTermSet(TermGroup group, string setName)
        {
            group.CreateTermSet(setName, Guid.NewGuid(), CultureInfo.CurrentCulture.LCID);
        }

        public TermGroup CreateTermGroup(string groupName)
        {
            var termStore = GetTermStore();
            return termStore.CreateGroup(groupName, Guid.NewGuid());
        }

        public TermGroup GetTermGroupByName(string groupName)
        {
            return GetTermStore().Groups.GetByName(groupName);
        }

        public TermStore GetTermStore()
        {
            TaxonomySession session = TaxonomySession.GetTaxonomySession(_context);
            return session.GetDefaultSiteCollectionTermStore();
        }

        public async Task SaveContextAsync()
        {
            await this._context.ExecuteQueryAsync();
        }

        public TermSet GetTermSetByName(string name, TermGroup group)
        {
            return group.TermSets.GetByName(name);
        }

        public void CreateTerm(string termName, TermSet terms)
        {
            terms.CreateTerm(termName, CultureInfo.CurrentCulture.LCID, Guid.NewGuid());
        }

        public void SetTermSetToTaxonomyField(ClientContext context, TermSet terms, string fieldName)
        {
            var field = context.Web.Fields.GetByInternalNameOrTitle(fieldName);
            TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(field);
            taxonomyField.SspId = terms.TermStore.Id;
            taxonomyField.TermSetId = terms.Id;
            taxonomyField.UpdateAndPushChanges(true);
        }

        public void SetFieldValueToTaxonomyField(ListItem item, TermCollection collection, ClientContext context, string fieldName)
        {
            var field = context.Web.Fields.GetByInternalNameOrTitle(fieldName);
            TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(field);
            taxonomyField.SetFieldValueByTermCollection(item, collection, CultureInfo.CurrentCulture.LCID);
            taxonomyField.UpdateAndPushChanges(true);
        }
    }
}

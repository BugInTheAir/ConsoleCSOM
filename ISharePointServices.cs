using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM
{
    public interface ISharePointServices
    {
        void CreateList(string name);
        TermGroup CreateTermGroup(string groupName);
        TermGroup GetTermGroupByName(string groupName);
        TermSet GetTermSetByName(string name, TermGroup group);
        TermStore GetTermStore();
        void CreateTermSet(TermGroup group, string setName);
        void CreateTerm(string termName, TermSet terms);
        Task SaveContextAsync();
        
    }
    public class SharePointServices : ISharePointServices
    {
        private readonly ClientContext _context;

        public SharePointServices(ClientContext context)
        {
            _context = context;
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
    }
}

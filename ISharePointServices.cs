using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Collections.Generic;
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
        void CreateDocumentList(string name);
        void SetTermSetToTaxonomyField(ClientContext context, TermSet terms, string fieldName);
        void SetFieldValueToTaxonomyField(ListItem item, TermCollection collection, ClientContext context, string fieldName);
        void CreateFolderInDocument(List documents, string folderInternalName, string folderName, string folderUrl);
    }
}

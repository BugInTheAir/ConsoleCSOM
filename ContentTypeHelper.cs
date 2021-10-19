using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM
{
    public class ContentTypeHelper
    {
        public async Task AddFieldToListContentType(List myRootList, FieldCollection sourceField, ClientContext ctx, IEnumerable<string> fieldNames, string contentTypeName, bool isContaintChildContentType)
        {
            ctx.Load(myRootList, e => e.ContentTypes, e => e.Fields);
            ctx.Load(sourceField);
            var contentCollection = myRootList.ContentTypes;
            await ctx.ExecuteQueryAsync();
            ContentType myContentType = null;
            foreach (var item in fieldNames)
            {
                myContentType = AddFieldToMyContentType(sourceField, contentCollection, item, contentTypeName, isContaintChildContentType);
            }
            if (myContentType == null)
                return;
            myRootList.Update();
            await ctx.ExecuteQueryAsync();
        }

        private static ContentType AddFieldToMyContentType(FieldCollection sourceFieldCollection, ContentTypeCollection contentCollection, string fieldName, string myContentTypeName, bool isContainChild)
        {
            var myContentType = contentCollection.Where(x => x.Name.Equals(myContentTypeName)).FirstOrDefault();
            FieldLinkCreationInformation info = new FieldLinkCreationInformation();
            info.Field = sourceFieldCollection.Where(x => x.Title.Equals(fieldName)).FirstOrDefault();
            myContentType.FieldLinks.Add(info);
            myContentType.Update(isContainChild);
            return myContentType;
        }

        public async Task SetMyContentTypeAsDefault(ClientContext ctx, string contentTypeNameToBeTop, List myListToBeSet)
        {
            var currentOrder = myListToBeSet.ContentTypes;
            ctx.Load(currentOrder);
            await ctx.ExecuteQueryAsync();
            IList<ContentTypeId> reverseOrder = new List<ContentTypeId>();
            foreach (var type in currentOrder)
            {
                if (type.Name.Equals(contentTypeNameToBeTop))
                {
                    reverseOrder.Add(type.Id);
                }
            }

            myListToBeSet.RootFolder.UniqueContentTypeOrder = reverseOrder;
            myListToBeSet.RootFolder.Update();
            myListToBeSet.Update();
        }

        public  async Task CreateCustomContentType(ClientContext ctx, string contentTypeName, string des, string group, string parentContentTypeName)
        {
            var contentCollection = ctx.Web.ContentTypes;
            ctx.Load(contentCollection);
            await ctx.ExecuteQueryAsync();
            var parentType = contentCollection.Where(x => x.Name.Equals(parentContentTypeName)).FirstOrDefault();
            if (parentType is null)
                return;
            ContentTypeCreationInformation oContentTypeCreationInformation = new ContentTypeCreationInformation();

            oContentTypeCreationInformation.Name = contentTypeName;

            oContentTypeCreationInformation.Description = des;

            oContentTypeCreationInformation.Group = group;

            oContentTypeCreationInformation.ParentContentType = parentType;

            ContentType oContentType = contentCollection.Add(oContentTypeCreationInformation);

        }
    }
}

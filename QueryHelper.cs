using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM
{
    public class QueryHelper
    {
        public static CamlQuery GetQueryByTitle(string titleName, string relativePath = null)
        {
            if(relativePath is null)
            {
                return new CamlQuery()
                {
                    ViewXml = @$"<View>
                                <Query>
                                    <Where>
                                         <Eq>
                                            <FieldRef Name=""Title""></FieldRef>
                                            <Value Type=""Text"">{titleName}</Value>
                                          </Eq>
                                  </Where>
                                </Query>
                            </View>"
                };
            }
            return new CamlQuery()
            {
                ViewXml = @$"<View>
                                <Query>
                                    <Where>
                                         <Eq>
                                            <FieldRef Name=""Title""></FieldRef>
                                            <Value Type=""Text"">{titleName}</Value>
                                          </Eq>
                                  </Where>
                                </Query>
                            </View>",
                FolderServerRelativeUrl = relativePath
            };
        }
    }
}

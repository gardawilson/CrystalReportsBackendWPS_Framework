using System;
using System.Reflection;

namespace CrystalReportsBackendWPS_Framework.Areas.HelpPage.ModelDescriptions
{
    public interface IModelDocumentationProvider
    {
        string GetDocumentation(MemberInfo member);

        string GetDocumentation(Type type);
    }
}
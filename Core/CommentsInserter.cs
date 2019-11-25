using EnvDTE;
using Microsoft.VisualStudio.VCCodeModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComI.Core
{
    class VCCommentsInserter
    {
        public static void InsertCommentForCodeElement(VCCodeElement codeElement, string snippetText)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            EditPoint elementStartPoint = codeElement.StartPointOf[vsCMPart.vsCMPartWholeWithAttributes, vsCMWhere.vsCMWhereDeclaration].CreateEditPoint();
            EditPoint elementEndPoint = codeElement.EndPointOf[vsCMPart.vsCMPartWholeWithAttributes, vsCMWhere.vsCMWhereDeclaration].CreateEditPoint();

            EditPoint insertPoint = elementStartPoint.CreateEditPoint();

            insertPoint.Insert("\n");
            InsertSingleComment(insertPoint, snippetText);
            elementStartPoint.SmartFormat(elementEndPoint);
        }

        private static void InsertSingleComment(EditPoint elementStartPoint, string snippetText)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            elementStartPoint.Insert("// " + snippetText + "\n");
        }
    }
}

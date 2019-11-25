using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.VCCodeModel;
using Task = System.Threading.Tasks.Task;

namespace ComI.Core
{
    class VCSelectionFinder
    {
        public static List<VCCodeElement> GetSelectedCodeElements(DTE2 projectModel)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            List<VCCodeElement> output = new List<VCCodeElement>();
            TextSelection selection = projectModel.ActiveDocument?.Selection as TextSelection;

            if (selection.TopPoint != null && selection.BottomPoint != null &&
                ((selection.TopPoint.Line != selection.BottomPoint.Line) ||
                (selection.TopPoint.DisplayColumn != selection.BottomPoint.DisplayColumn)))
            {
                output = GetMultiLineSelection(projectModel.ActiveDocument, selection);
            }
            else if (selection.ActivePoint != null)
                output.Add( GetSingleLineSelection(projectModel.ActiveDocument, selection.ActivePoint) );

            return output;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="codeElements"></param>
        /// <param name="selectionTopPoint"></param>
        /// <param name="selectionBottomPoint"></param>
        /// <returns></returns>
        private static List<VCCodeElement> GetSelectedItemsBetweenPoints(CodeElements codeElements, EditPoint selectionTopPoint, EditPoint selectionBottomPoint)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            List<VCCodeElement> output = new List<VCCodeElement>();

            for (int i = 1; i <= codeElements.Count; i++)
            {
                if (codeElements.Item(i) is VCCodeElement codeElement)
                {
                    TextPoint elementStartPoint = codeElement.StartPointOf[vsCMPart.vsCMPartWholeWithAttributes, vsCMWhere.vsCMWhereDeclaration];
                    TextPoint elementEndPoint = codeElement.EndPointOf[vsCMPart.vsCMPartWholeWithAttributes, vsCMWhere.vsCMWhereDeclaration];
                    EditPoint elementNameEndPoint = FindNameEnd(elementStartPoint.CreateEditPoint()).CreateEditPoint();

                    if (IsSelectionContaining(selectionTopPoint, selectionBottomPoint, elementStartPoint, elementEndPoint))
                    {
                        if (IsSelectionContaining(selectionTopPoint, selectionBottomPoint, elementStartPoint, elementNameEndPoint))
                        {
                            output.Add(codeElement);
                        }
                        output.AddRange(GetSelectedItemsBetweenPoints(codeElement.Children, selectionTopPoint, selectionBottomPoint));
                    }
                }
            }
            return output;
        }
       
        /// <summary>
        /// 
        /// </summary>
        /// <param name="selectionTop"></param>
        /// <param name="selectionBottom"></param>
        /// <param name="startPoint"></param>
        /// <param name="endPoint"></param>
        /// <returns></returns>
        private static bool IsSelectionContaining(EditPoint selectionTop, EditPoint selectionBottom, TextPoint startPoint, TextPoint endPoint)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return !(startPoint.GreaterThan(selectionBottom) || !endPoint.GreaterThan(selectionTop));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="elementStartPoint"></param>
        /// <returns></returns>
        private static TextPoint FindNameEnd(EditPoint elementStartPoint)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            EditPoint2 actualPoint = (EditPoint2)elementStartPoint.CreateEditPoint();
            string text = actualPoint.GetText(1);
            while (text != "{" && text != ";")
            {
                actualPoint.CharRight();
                text = actualPoint.GetText(1);
            }
            return actualPoint.CreateEditPoint();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="activeDocuemnt"></param>
        /// <param name="selection"></param>
        /// <returns></returns>
        private static List<VCCodeElement> GetMultiLineSelection(Document activeDocuemnt, TextSelection selection)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            List<VCCodeElement> output;

            EditPoint selectionTopPoint = selection.TopPoint.CreateEditPoint();
            EditPoint selectionBottomPoint = selection.BottomPoint.CreateEditPoint();

            CodeElements codeElements = activeDocuemnt.ProjectItem.FileCodeModel.CodeElements;
            output = GetSelectedItemsBetweenPoints(codeElements, selectionTopPoint, selectionBottomPoint);

            return output;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="activeDocuemnt"></param>
        /// <param name="activePoint"></param>
        /// <returns></returns>
        private static VCCodeElement GetSingleLineSelection(Document activeDocuemnt, TextPoint activePoint)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return activeDocuemnt?.ProjectItem?.FileCodeModel?.CodeElementFromPoint(activePoint, vsCMElement.vsCMElementVCBase) as VCCodeElement;
        }
    }
}

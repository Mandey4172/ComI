using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.VCCodeModel;
using Task = System.Threading.Tasks.Task;

namespace ComI
{
    /// <summary>
    /// DocumentCommand is used by the IDE to run on the code elements 
    /// (namespace, classes, functions, etc.) documentation functionality.
    /// </summary>
    internal sealed class DocumentCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("c2b40c43-4a49-4ddb-b45e-1b44dfec7690");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private DocumentCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static DocumentCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in DocumentCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new DocumentCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private async void Execute(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            await DocumentAsync();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private async Task DocumentAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            if (await ServiceProvider.GetServiceAsync(typeof(DTE)) is DTE2 projectModel)
            {
                TextSelection selection = projectModel.ActiveDocument?.Selection as TextSelection;

                if (selection.TopPoint != null && selection.BottomPoint != null &&
                    ((selection.TopPoint.Line != selection.BottomPoint.Line) ||
                    (selection.TopPoint.DisplayColumn != selection.BottomPoint.DisplayColumn)))
                {
                    List<VCCodeElement> selectetElements = MultiLineSelection(projectModel.ActiveDocument, selection);
                    foreach (CodeElement element in selectetElements)
                    {
                        Debug.WriteLine(element.FullName);
                    }
                }
                else if (selection.ActivePoint != null)
                    Debug.WriteLine(SingleLineSelection(projectModel.ActiveDocument, selection.ActivePoint).Name);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="activeDocuemnt"></param>
        /// <param name="selection"></param>
        /// <returns></returns>
        private List<VCCodeElement> MultiLineSelection(Document activeDocuemnt, TextSelection selection)
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
        /// <param name="codeElements"></param>
        /// <param name="selectionTopPoint"></param>
        /// <param name="selectionBottomPoint"></param>
        /// <returns></returns>
        private List<VCCodeElement> GetSelectedItemsBetweenPoints(CodeElements codeElements, EditPoint selectionTopPoint, EditPoint selectionBottomPoint)
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

                    Debug.WriteLine("codeElements = " + codeElement.FullName);
                    Debug.WriteLine("elementStartPoint = " + elementStartPoint.AbsoluteCharOffset);
                    Debug.WriteLine("elementEndPoint = " + elementEndPoint.AbsoluteCharOffset);
                    //Debug.WriteLine("elementStartPointHeader = " + elementStartPointHeader.AbsoluteCharOffset);
                    Debug.WriteLine("elementNameEndPoint = " + elementNameEndPoint.AbsoluteCharOffset);
                    Debug.WriteLine("selectionTopPoint = " + selectionTopPoint.AbsoluteCharOffset);
                    Debug.WriteLine("selectionBottomPoint = " + selectionBottomPoint.AbsoluteCharOffset);
                    Debug.WriteLine("");

                    if (IsSelectionContaining(selectionTopPoint, selectionBottomPoint, elementStartPoint, elementEndPoint))
                    {
                        if (IsSelectionContaining(selectionTopPoint, selectionBottomPoint, elementStartPoint, elementNameEndPoint))
                        {
                            Debug.WriteLine("Added -> codeElements = " + codeElement.FullName);
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
        private bool IsSelectionContaining(EditPoint selectionTop, EditPoint selectionBottom, TextPoint startPoint, TextPoint endPoint)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return !(startPoint.GreaterThan(selectionBottom) || !endPoint.GreaterThan(selectionTop));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="elementStartPoint"></param>
        /// <returns></returns>
        private TextPoint FindNameEnd(EditPoint elementStartPoint)
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
        /// <param name="activePoint"></param>
        /// <returns></returns>
        private CodeElement SingleLineSelection(Document activeDocuemnt, TextPoint activePoint)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return activeDocuemnt?.ProjectItem?.FileCodeModel?.CodeElementFromPoint(activePoint, vsCMElement.vsCMElementVCBase) as CodeElement;
        }
    }
}

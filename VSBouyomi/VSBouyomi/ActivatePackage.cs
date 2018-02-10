using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace VSBouyomi {
	/// <summary>
	/// This is the class that implements the package exposed by this assembly.
	/// </summary>
	/// <remarks>
	/// <para>
	/// The minimum requirement for a class to be considered a valid package for Visual Studio
	/// is to implement the IVsPackage interface and register itself with the shell.
	/// This package uses the helper classes defined inside the Managed Package Framework (MPF)
	/// to do it: it derives from the Package class that provides the implementation of the
	/// IVsPackage interface and uses the registration attributes defined in the framework to
	/// register itself and its components with the shell. These attributes tell the pkgdef creation
	/// utility what data to put into .pkgdef file.
	/// </para>
	/// <para>
	/// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
	/// </para>
	/// </remarks>
	[PackageRegistration(UseManagedResourcesOnly = true)]
	[InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
	[ProvideMenuResource("Menus.ctmenu", 1)]
	[Guid(ActivatePackage.PackageGuidString)]
	[SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
	[ProvideAutoLoad(UIContextGuids.NoSolution)]
	[ProvideAutoLoad(UIContextGuids.SolutionExists)]
	[ProvideAutoLoad(UIContextGuids.CodeWindow)]
	[ProvideAutoLoad(UIContextGuids.EmptySolution)]
	[ProvideOptionPage(typeof(OptionDialog),"VSBouyomi","General",0,0,true)]
	public sealed class ActivatePackage : Package {
		/// <summary>
		/// ActivatePackage GUID string.
		/// </summary>
		public const string PackageGuidString = "fc28ac53-95c3-4339-b941-00c52bb18882";

		/// <summary>
		/// Initializes a new instance of the <see cref="Activate"/> class.
		/// </summary>
		public ActivatePackage() {
			// Inside this method you can place any initialization code that does not require
			// any Visual Studio service because at this point the package object is created but
			// not sited yet inside Visual Studio environment. The place to do all the other
			// initialization is the Initialize method.
		}

		#region Package Members

		public EventHolder EventHolder { get; set; }

		/// <summary>
		/// Initialization of the package; this method is called right after the package is sited, so this is the place
		/// where you can put all the initialization code that rely on services provided by VisualStudio.
		/// </summary>
		protected override void Initialize() {

			var dte = (DTE)GetService(typeof(DTE));
			this.EventHolder = new EventHolder
			{
				BuildEvents = dte.Events.BuildEvents,
				SolutionEvents = dte.Events.SolutionEvents,
				DebuggerEvents = dte.Events.DebuggerEvents,
				DteEvents = dte.Events.DTEEvents
			};

			Activate.Initialize(this,this.EventHolder);
			base.Initialize();
		}

		public static string ConfigXmlPath {
			get
			{
				var codebase = typeof(ActivatePackage).Assembly.CodeBase;
				var uri = new Uri(codebase, UriKind.Absolute);
				return System.IO.Path.GetDirectoryName(uri.LocalPath) + "\\Resources\\config.xml";
			}
		}

		#endregion
	}
}

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using System.Windows.Forms;
using System.Diagnostics;

namespace VSBouyomi {
	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed partial class Activate {
		/// <summary>
		/// Command ID.
		/// </summary>
		public const int CommandId = 0x0100;

		/// <summary>
		/// Command menu group (command set GUID).
		/// </summary>
		public static readonly Guid CommandSet = new Guid("066be89d-e710-4f88-8e72-dc1ff0ab0b10");

		/// <summary>
		/// VS Package that provides this command, not null.
		/// </summary>
		private readonly Package package;

		/// <summary>
		/// Initializes a new instance of the <see cref="Activate"/> class.
		/// Adds our command handlers for menu (commands must exist in the command table file)
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		private Activate(Package package, EventHolder eventHolder) {
			this.package = package ?? throw new ArgumentNullException("package");

			this.EventHolder = eventHolder;
			Eventer = new Eventer(EventHolder, ActivatePackage.ConfigXmlPath);

			if (this.ServiceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService) {
				var menuCommandID = new CommandID(CommandSet, CommandId);
				var menuItem = new OleMenuCommand(this.MenuItemCallback, menuCommandID);
				commandService.AddCommand(menuItem);
			}
		}

		public EventHolder EventHolder { get; set; }
		public static Eventer Eventer { get; set; }

		/// <summary>
		/// Gets the instance of the command.
		/// </summary>
		public static Activate Instance {
			get;
			private set;
		}

		/// <summary>
		/// Gets the service provider from the owner package.
		/// </summary>
		private IServiceProvider ServiceProvider => this.package;

		/// <summary>
		/// Initializes the singleton instance of the command.
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		public static void Initialize(Package package, EventHolder eventHolder) {
			Instance = new Activate(package, eventHolder);
		}

		/// <summary>
		/// This function is the callback used to execute the command when the menu item is clicked.
		/// See the constructor to see how the menu item is associated with this function using
		/// OleMenuCommandService service and MenuCommand class.
		/// </summary>
		/// <param name="sender">SupportingEvents sender.</param>
		/// <param name="e">SupportingEvents args.</param>
		private void MenuItemCallback(object sender, EventArgs e) {
			var mycmd = sender as OleMenuCommand;
			if (mycmd != null) mycmd.Checked = !mycmd.Checked;


			if (mycmd.Checked) {
				Eventer.Activate();
			} else {
				Eventer.Deactivate();
			}
		
		}

		

	}
}
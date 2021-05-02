using System;
using System.Collections.Generic;
using System.Linq;
using asu_docx_validator.Validation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ElectronNET.API;
using ElectronNET.API.Entities;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace asu_docx_validator
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddRazorPages();
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseExceptionHandler("/Error");
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }

            app.UseHttpsRedirection();
            app.UseStaticFiles();

            app.UseRouting();

            app.UseAuthorization();

            app.UseEndpoints(endpoints => { endpoints.MapRazorPages(); });

            if (HybridSupport.IsElectronActive)
            {
                CreateWindow();
            }
        }

        private void CreateMenu()
        {
            MenuItem[] menu = null;

            MenuItem[] fileMenu = new MenuItem[]
            {
                new MenuItem
                {
                    Label = "Load file", Type = MenuType.normal, Click = async () =>
                    {
                        var mainWindow = Electron.WindowManager.BrowserWindows.First();
                        var options = new OpenDialogOptions()
                        {
                            Properties = new OpenDialogProperty[] {OpenDialogProperty.openFile},
                            Filters = new FileFilter[]
                            {
                                new FileFilter {Name = "Word Documents (.docx)", Extensions = new string[] {"docx"}}
                            }
                        };
                        string[] filePaths = await Electron.Dialog.ShowOpenDialogAsync(mainWindow, options);
                        WordprocessingDocument wordProcessingDocument =
                            WordprocessingDocument.Open(filePaths[0], false);
                        if (wordProcessingDocument.MainDocumentPart != null)
                        {
                            Document document = wordProcessingDocument.MainDocumentPart.Document;
                            Dictionary<String, List<String>> errorsMap = new Dictionary<string, List<string>>();
                            TitleValidator.Validate(errorsMap, document, filePaths[0]);
                            foreach (var error in errorsMap.Values)
                            {
                                Array.ForEach(error.ToArray(),Console.WriteLine);
                            }
                            validatePageMargins(document);
                        }
                    }
                },
                new MenuItem {Type = MenuType.separator},
                new MenuItem {Role = MenuRole.quit}
            };

            MenuItem[] viewMenu = new MenuItem[]
            {
                new MenuItem {Role = MenuRole.reload},
                new MenuItem {Role = MenuRole.forcereload},
                new MenuItem {Role = MenuRole.toggledevtools},
                new MenuItem {Type = MenuType.separator},
                new MenuItem {Role = MenuRole.resetzoom},
                new MenuItem {Role = MenuRole.zoomin},
                new MenuItem {Role = MenuRole.zoomout},
                new MenuItem {Type = MenuType.separator},
                new MenuItem {Role = MenuRole.togglefullscreen}
            };

            menu = new MenuItem[]
            {
                new MenuItem {Label = "File", Type = MenuType.submenu, Submenu = fileMenu},
                new MenuItem {Label = "View", Type = MenuType.submenu, Submenu = viewMenu}
            };

            Electron.Menu.SetApplicationMenu(menu);
        }

        private void validatePageMargins(Document document)
        {
            IEnumerable<PageMargin> pageMarginEnumerable = document.Descendants<PageMargin>();
            PageMargin pageMargin = pageMarginEnumerable.ToList()[0];
            if (pageMargin.Top != 1418 || pageMargin.Bottom != 1418 || pageMargin.Left != 1418 ||
                pageMargin.Right != 1418)
            {
                Electron.Notification.Show(new NotificationOptions("Document Margins",
                    "your document margins are not correct"));
            }
        }

        private async void CreateWindow()
        {
            CreateMenu();
            var window = await Electron.WindowManager.CreateWindowAsync();
            window.OnClosed += () => { Electron.App.Quit(); };
        }
    }
}
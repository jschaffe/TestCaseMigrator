using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

/// <summary>
/// NOTE: Proper Area and Iteration Paths must be created in the target project prior to running this code.
/// </summary>
namespace TestCaseMigrator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private TfsTeamProjectCollection tpSourceCollection;
        private TfsTeamProjectCollection tpTargetCollection;
        private string sourceProjectName;
        private string targetProjectName;
        private string targetUri;
        private ITestManagementService tmSourceService;
        private ITestManagementTeamProject tmSourceProject;
        private ITestManagementService tmTargetService;
        private ITestManagementTeamProject tmTargetProject;
        private WorkItemStore wisTargetProject;

        private Dictionary<int, int> sharedStepMapping = new Dictionary<int, int>();
        private Dictionary<string, string> userMapping = new Dictionary<string, string>();

        private readonly SynchronizationContext synchronizationContext;

        public MainWindow()
        {
            InitializeComponent();

            synchronizationContext = SynchronizationContext.Current;

            InitUserMapping();
        }

        /// <summary>
        /// Establish connection to source TFS team project.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSource_Click(object sender, RoutedEventArgs e)
        {
            TeamProjectPicker projectPicker = new TeamProjectPicker();

            projectPicker.ShowDialog();

            if (projectPicker.SelectedTeamProjectCollection != null)
            {
                tpSourceCollection = projectPicker.SelectedTeamProjectCollection;
                targetUri = $"{tpSourceCollection.Uri}//{projectPicker.SelectedProjects[0].Name}/";
                txtSourceProject.Text = $"{tpSourceCollection}{Environment.NewLine}{projectPicker.SelectedProjects[0]}";
                tmSourceService = tpSourceCollection.GetService<ITestManagementService>();
                sourceProjectName = projectPicker.SelectedProjects[0].Name;
                tmSourceProject = tmSourceService.GetTeamProject(sourceProjectName);
            }
        }

        /// <summary>
        /// Establish connection to target VSTS team project
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTarget_Click(object sender, RoutedEventArgs e)
        {
            // Establish connection to target system.
            TeamProjectPicker projectPicker = new TeamProjectPicker();

            projectPicker.ShowDialog();

            if (projectPicker.SelectedTeamProjectCollection != null)
            {
                tpTargetCollection = projectPicker.SelectedTeamProjectCollection;
                txtTargetProject.Text = String.Format(@"{0}{1}{2}", tpTargetCollection.ToString(), Environment.NewLine, projectPicker.SelectedProjects[0].ToString());
                tmTargetService = tpTargetCollection.GetService<ITestManagementService>();
                targetProjectName = projectPicker.SelectedProjects[0].Name;
                tmTargetProject = tmTargetService.GetTeamProject(targetProjectName);

                wisTargetProject = new WorkItemStore(tpTargetCollection, WorkItemStoreFlags.BypassRules);
            }
        }

        private void btnMigrate_Click(object sender, RoutedEventArgs e)
        {
            PerformMigration();
        }
        
        /// <summary>
        /// Main method to initialize async execution of migration
        /// </summary>
        private async void PerformMigration()
        {
            var testCaseTask = new Task(() => { MigrateTestCases();});

            var t = Task.Run(() => {
                MigrateSharedSteps();
            }).ContinueWith((antecedant) =>
            {
                testCaseTask.Start();
            });
        }
        
        /// <summary>
        /// Updates the status area of the main window
        /// </summary>
        /// <param name="message">The message to display in the status field.</param>
        public void UpdateStatus(object message)
        {
            txtStatus.Text = message as string;
        }

        /// <summary>
        /// Migrate the Shared Steps from source to target project
        /// </summary>
        /// <remarks>
        /// Shared Step mapping (source ID to target ID) will be maintained in memory to assign the appropriate shared step ID
        /// to any referencing test cases during their migration.
        /// </remarks>
        private void MigrateSharedSteps()
        {
            int sharedStepCount = 0;
            int sharedStepFailures = 0;
            Trace.WriteLine($"** Starting Shared Step Migration: {DateTime.Now.ToLongTimeString()}");
            Stopwatch sw = new Stopwatch();

            sw.Start();
            var sourceSharedStepQuery = $"SELECT * FROM WorkItems WHERE [Work Item Type] = 'Shared Steps' AND [Team Project] = '{sourceProjectName}'";
            IEnumerable<ISharedStep> sharedSteps = tmSourceProject.SharedSteps.Query(sourceSharedStepQuery);
            foreach (var sourceSharedStep in sharedSteps)
            {
                try
                {
                    sharedStepCount++;
                    synchronizationContext.Post(UpdateStatus, $"Processing shared step {sharedStepCount} ({sharedStepFailures})");

                    Trace.WriteLine($"Shared Step {sourceSharedStep.Id}: {sourceSharedStep.Title}");
                    var targetSharedStep = new WorkItem(wisTargetProject.Projects[targetProjectName].WorkItemTypes["Shared Steps"]);

                    targetSharedStep.Title = sourceSharedStep.Title;
                    targetSharedStep.Description = sourceSharedStep.Description;
                    targetSharedStep["Priority"] = sourceSharedStep.Priority;
                    targetSharedStep.IterationPath = sourceSharedStep.WorkItem.IterationPath.Replace(sourceProjectName, targetProjectName);
                    targetSharedStep.AreaPath = sourceSharedStep.WorkItem.AreaPath.Replace(sourceProjectName, targetProjectName);
                    targetSharedStep.State = sourceSharedStep.State;
                    targetSharedStep[CoreField.AssignedTo] = GetUser(sourceSharedStep.WorkItem[CoreField.AssignedTo] as string);
                    targetSharedStep.Tags = sourceSharedStep.WorkItem.Tags;
                    targetSharedStep.Save();

                    var targetSharedStepQuery = $"SELECT * FROM WorkItems WHERE [Work Item Type] = 'Shared Steps' AND [Id] = {targetSharedStep.Id}";
                    var sharedStep = tmTargetProject.SharedSteps.Query(targetSharedStepQuery).FirstOrDefault();

                    foreach (var action in sourceSharedStep.Actions)
                    {
                        ITestStep step = action as ITestStep;
                        if (step != null)
                        {
                            Trace.WriteLine($"\t\tStep {step.Id}: Title = {step.Title}, ExpectedResult = {step.ExpectedResult}");
                            var targetTestStep = sharedStep.CreateTestStep();
                            targetTestStep.Title = step.Title;
                            targetTestStep.ExpectedResult = step.ExpectedResult;

                            sharedStep.Actions.Add(targetTestStep);
                        }
                    }

                    sharedStep.Save();
                    sharedStepMapping.Add(sourceSharedStep.Id, targetSharedStep.Id);
                }
                catch (Exception ex)
                {
                    sharedStepFailures++;
                    Trace.WriteLine($"Error processing shared step ({sourceSharedStep.Id}:{sourceSharedStep.Title}): {ex.Message}");
                }
            }
            sw.Stop();
            Trace.WriteLine($"** Shared Step Migration Complete: {DateTime.Now.ToLongTimeString()}");
            Trace.WriteLine($"** Successfully migrated {sharedStepCount - sharedStepFailures} of {sharedStepCount} shared steps");
            Trace.WriteLine($"** Execution time {sw.Elapsed.TotalSeconds} seconds");
        }

        /// <summary>
        /// Migrate Test Cases from source to target project
        /// </summary>
        private void MigrateTestCases()
        {
            int testCaseCount = 0;
            int testCaseFailures = 0;
            Trace.WriteLine($"** Starting Test Case Migration: {DateTime.Now.ToLongTimeString()}");
            Stopwatch sw = new Stopwatch();

            sw.Start();
            var testCaseQuery = $"SELECT * FROM WorkItems WHERE [Work Item Type] = 'Test Case' AND [Team Project] = '{sourceProjectName}'";
            IEnumerable<ITestCase> testCases = tmSourceProject.TestCases.Query(testCaseQuery);
            foreach (var sourceTestCase in testCases)
            {
                try
                {
                    testCaseCount++;
                    synchronizationContext.Post(UpdateStatus, $"Processing test case {testCaseCount} ({testCaseFailures})");

                    Trace.WriteLine(String.Format("Test Case {0}: {1}", sourceTestCase.Id, sourceTestCase.Title));

                    var targetTestCase = new WorkItem(wisTargetProject.Projects[targetProjectName].WorkItemTypes["Test Case"]);

                    targetTestCase.Title = sourceTestCase.Title;
                    targetTestCase.Description = sourceTestCase.Description;
                    targetTestCase["Priority"] = sourceTestCase.Priority;
                    targetTestCase.IterationPath = sourceTestCase.WorkItem.IterationPath.Replace(sourceProjectName, targetProjectName);
                    targetTestCase.AreaPath = sourceTestCase.WorkItem.AreaPath.Replace(sourceProjectName, targetProjectName);
                    targetTestCase.State = sourceTestCase.State;
                    targetTestCase[CoreField.AssignedTo] = GetUser(sourceTestCase.WorkItem[CoreField.AssignedTo] as string);
                    targetTestCase.Tags = sourceTestCase.WorkItem.Tags;
                    // TODO: Create ReflectedWorkitemId field on TestCase work item type in VSTS
                    targetTestCase["ReflectedWorkitemId"] = targetUri + sourceTestCase.Id;
                    // TODO: Add any custom fields here.
                    //targetTestCase["Precondition"] = sourceTestCase.CustomFields["Pre Condition"].Value;
                    targetTestCase.Save();

                    var targetTestCaseQuery = $"SELECT * FROM WorkItems WHERE [Work Item Type] = 'Test Case' AND [Id] = {targetTestCase.Id}";
                    var testCase = tmTargetProject.TestCases.Query(targetTestCaseQuery).FirstOrDefault();
                    testCase.WorkItem.Open();

                    foreach (var action in sourceTestCase.Actions)
                    {
                        ITestStep step = action as ITestStep;
                        if (step != null)
                        {
                            Trace.WriteLine($"\tStep {step.Id}: Title = {step.Title}, ExpectedResult = {step.ExpectedResult}");
                            var targetTestStep = testCase.CreateTestStep();
                            targetTestStep.Title = step.Title;
                            targetTestStep.ExpectedResult = step.ExpectedResult;
                            testCase.Actions.Add(targetTestStep);
                        }
                        else
                        {
                            // Since this was not a regular test step, assume Shared Step
                            ISharedStepReference sourceSharedStep = action as ISharedStepReference;
                            if (sourceSharedStep != null)
                            {
                                Trace.WriteLine(String.Format("\tShared Step Reference: {0}", sourceSharedStep.Id));
                                var targetSharedStep = testCase.CreateSharedStepReference();
                                int mappedId = 0;
                                sharedStepMapping.TryGetValue(sourceSharedStep.SharedStepId, out mappedId);
                                if (mappedId != 0)
                                {
                                    targetSharedStep.SharedStepId = mappedId;
                                    testCase.Actions.Add(targetSharedStep);
                                }
                                else
                                {
                                    var targetTestStep = testCase.CreateTestStep();
                                    targetTestStep.Title = $"PLACEHOLDER: Shared Step (original ID:{sourceSharedStep.SharedStepId}";
                                    testCase.Actions.Add(targetTestStep);
                                }
                            }
                        }
                    }

                    testCase.Save();
                }
                catch (Exception ex)
                {
                    testCaseFailures++;
                    Trace.WriteLine($"Error processing test case ({sourceTestCase.Id}:{sourceTestCase.Title}): {ex.Message}");
                }
            }
            sw.Stop();
            Trace.WriteLine($"** Test Case Migration Complete: {DateTime.Now.ToLongTimeString()}");
            Trace.WriteLine($"** Successfully migrated {testCaseCount - testCaseFailures} of {testCaseCount} test cases");
            Trace.WriteLine($"** Execution time {sw.Elapsed.TotalSeconds} seconds");

            synchronizationContext.Post(UpdateStatus, $"Processing test case {testCaseCount} ({testCaseFailures}). PROCESSING COMPLETE.");
        }

        /// <summary>
        /// Maps user display names from source to target for updateing the AssignedTo field.  If the user's display name is the same
        /// in the source and target, no mapping is required.
        /// </summary>
        private void InitUserMapping()
        {
            // TODO: Provide any display name user mappings here to correctly update AssignedTo field in target project.
            userMapping.Add("James Schaffer", "Schaffer, James");
        }

        /// <summary>
        /// Given the assignedTo value from the source project, return the corresponding display name in target project
        /// </summary>
        /// <param name="assignedTo">The display name for the AssignedTo field from the source project</param>
        /// <returns>
        /// The corresponding display name for the target project is a mapping is provided.  Otherwise return the original 
        /// assignedTo value.
        /// </returns>
        private string GetUser(string assignedTo)
        {
            if (!string.IsNullOrEmpty(assignedTo))
            {
                string mappedUser;
                userMapping.TryGetValue(assignedTo, out mappedUser);
                if (!String.IsNullOrEmpty(mappedUser))
                {
                    return mappedUser;
                }
            }

            return assignedTo;
        }
    }
}

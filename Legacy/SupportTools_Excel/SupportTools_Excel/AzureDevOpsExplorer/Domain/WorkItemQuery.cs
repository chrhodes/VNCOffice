using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;

using VNC;
using VNC.Core;

namespace SupportTools_Excel.AzureDevOpsExplorer.Domain
{
    public class WorkItemQuery
    {
        #region Properties

        public string Name { get; set; }
        // NOTE(crhodes)
        // We can set a default here or do in ViewModel PopulateWorkItemQueries
        public string QueryWithTokens { get; set; }

        public string Query { get; set; }

        public List<string> Fields { get; set; }

        //public Boolean RetrieveRevisions { get; set; }

        #endregion

        #region Public Methods

        /// <summary>
        /// Maybe this should be PrepareQuery
        /// </summary>
        /// <param name="options"></param>
        /// <param name="projectName"></param>
        public void ReplaceQueryTokens(
            Options_AZDO_TFS options,
            string projectName = null)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            Query = QueryWithTokens;

            //if ((options.WorkItemQuerySpec.Fields?.Count ?? 0) > 0)
            //{
                Query = Query.Replace("@FIELDS", GetFieldsToRetrieve(options));
            //}
            //else
            //{
            //    Query = Query.Replace("@FIELDS", "[System.Id]");
            //}

            // TODO(crhodes)
            // Until we better think through how to handle looping delays across Projects
            // Support both the @PROJECT token and TeamProjectFilter

            //if ((options.TeamProjects?.Count ?? 0) > 0)
            //{
            //    Query += GetTeamProjectsFilter(options);
            //}
            //else if (projectName != null)
            //{
            //    Query = Query.Replace("@PROJECT", $"{projectName}");
            //}

            if (Query.Contains("@PROJECT") && projectName != null)
            {
                Query = Query.Replace("@PROJECT", $"{projectName}");
            }
            else
            {
                if ((options.TeamProjects?.Count ?? 0) > 0)
                {
                    Query += GetTeamProjectsFilter(options);
                }
            }

            Query = Query.Replace("@STARTDATE", options.StartDate.ToShortDateString());
            Query = Query.Replace("@ENDDATE", options.EndDate.ToShortDateString());

            if ((options.WorkItemTypes?.Count ?? 0) > 0)
            {
                Query += GetWorkItemTypesFilter(options);
            }

            Log.APPLICATION($"Exit ({Query})", Common.LOG_CATEGORY, startTicks);
        }

        private string GetFieldsToRetrieve(Options_AZDO_TFS options)
        {
            StringBuilder fields = new StringBuilder("[System.Id]");

            if ((options.WorkItemQuerySpec.Fields?.Count ?? 0) > 0)
            {
                foreach (string field in options.WorkItemQuerySpec.Fields)
                {
                    fields.Append($", [{field}]");
                }

            }

            return fields.ToString();
        }


        private string GetWorkItemTypesFilter(Options_AZDO_TFS options)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            StringBuilder filter = new StringBuilder();

            if (options.WorkItemTypes.Count == 1)
            {
                filter.Append(" AND [System.WorkItemType] == " + $"{ options.WorkItemTypes[0].WrapInSngQuotes() }");
            }
            else
            {
                filter.Append(" AND [System.WorkItemType] in (");

                if (options.WorkItemTypes.Count >= 1)
                {
                    filter.Append('\'').Append(options.WorkItemTypes[0]).Append('\'');
                }

                if (options.WorkItemTypes.Count > 1)
                {
                    for (int i = 1; i < options.WorkItemTypes.Count; i++)
                    {
                        filter.Append(", '").Append(options.WorkItemTypes[i]).Append('\'');
                    }
                }

                filter.Append(')');
            }

            Log.APPLICATION($"Exit ({filter})", Common.LOG_CATEGORY, startTicks);

            return filter.ToString();

        }

        private string GetTeamProjectsFilter(Options_AZDO_TFS options)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            StringBuilder filter = new StringBuilder();

            if (options.TeamProjects.Count == 1)
            {
                filter.Append( " AND [System.TeamProject] == " + $"{ options.TeamProjects[0].WrapInSngQuotes() }");
            }
            else
            {
                filter.Append( " AND [System.TeamProject] in (");

                if (options.TeamProjects.Count >= 1)
                {
                    filter.Append('\'').Append(options.TeamProjects[0]).Append('\'');
                }

                if (options.TeamProjects.Count > 1)
                {
                    for (int i = 1; i < options.TeamProjects.Count; i++)
                    {
                        //filter.Append($", '{options.TeamProjects[i]}'");
                        filter.Append(", '").Append(options.TeamProjects[i]).Append('\'');
                    }
                }

                filter.Append(')');
            }

            Log.APPLICATION($"Exit ({filter})", Common.LOG_CATEGORY, startTicks);

            return filter.ToString();
        }

        #endregion
    }
}

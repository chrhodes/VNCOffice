using System.Windows;

using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;

using SupportTools_Visio.Domain;

namespace SupportTools_Visio.Actions
{
    public class AZDOPageLayout
    {
        internal static Point CalculateInsertionPointLinkedWorkItems(Point initialPosition, 
            WorkItem linkedWorkItem, WorkItemShapeInfo activeShape, WorkItemOffsets workItemOffsets)
        {
            Point newInsertionPoint = new Point();

            double height = activeShape.Height;
            double width = activeShape.Width;

            string shapeWorkItemType = activeShape.WorkItemType;

            switch (linkedWorkItem.Fields["System.WorkItemType"])
            {
                case "Bug":
                    switch (shapeWorkItemType)
                    {
                        case "Bug":
                            //workItemOffsets.Bug.DecrementHorizontal(width);
                            workItemOffsets.Bug.DecrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Bug.X;
                            newInsertionPoint.Y = workItemOffsets.Bug.Y;
                            break;

                        case "Epic":
                            //workItemOffsets.Bug.DecrementHorizontal(width);
                            workItemOffsets.Bug.DecrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Bug.X;
                            newInsertionPoint.Y = workItemOffsets.Bug.Y;
                            break;

                        case "Feature":
                            //workItemOffsets.Bug.DecrementHorizontal(width);
                            workItemOffsets.Bug.DecrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Bug.X;
                            newInsertionPoint.Y = workItemOffsets.Bug.Y;
                            break;

                        case "Release":
                            //workItemOffsets.Bug.DecrementHorizontal(width);
                            workItemOffsets.Bug.DecrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Bug.X;
                            newInsertionPoint.Y = workItemOffsets.Bug.Y;
                            break;

                        case "Requirement":
                            workItemOffsets.Bug.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Bug.X;
                            newInsertionPoint.Y = workItemOffsets.Bug.Y;
                            break;

                        case "Task":
                            //workItemOffsets.Bug.DecrementHorizontal(width);
                            workItemOffsets.Bug.DecrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Bug.X;
                            newInsertionPoint.Y = workItemOffsets.Bug.Y;
                            break;

                        case "Test Case":
                            if (workItemOffsets.Release.Count > 0)
                            {
                                workItemOffsets.UserStory.DecrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.Release.X;
                                newInsertionPoint.Y = workItemOffsets.Release.Y;
                            }
                            else if (workItemOffsets.UserStory.Count > 0)
                            {
                                workItemOffsets.UserStory.DecrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }
                            else if (workItemOffsets.Bug.Count > 0)
                            {
                                workItemOffsets.Bug.DecrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.Bug.X;
                                newInsertionPoint.Y = workItemOffsets.Bug.Y;
                            }
                            else
                            {
                                workItemOffsets.Unknown.IncrementHorizontal(width);
                                newInsertionPoint.X = workItemOffsets.Unknown.X;
                                newInsertionPoint.Y = workItemOffsets.Unknown.Y;
                            }

                            break;

                        case "User Needs":
                            workItemOffsets.Bug.DecrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Bug.X;
                            newInsertionPoint.Y = workItemOffsets.Bug.Y;
                            break;

                        case "User Story":
                            workItemOffsets.Bug.DecrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Bug.X;
                            newInsertionPoint.Y = workItemOffsets.Bug.Y;
                            break;

                        default:
                            // TODO(crhodes)
                            // What should this do???
                            break;
                    }

                    break;

                case "Epic":
                    switch (shapeWorkItemType)
                    {
                        case "Bug":
                            workItemOffsets.Epic.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Epic.X;
                            newInsertionPoint.Y = workItemOffsets.Epic.Y;
                            break;

                        case "Epic":
                            workItemOffsets.Epic.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Epic.X;
                            newInsertionPoint.Y = workItemOffsets.Epic.Y;
                            break;

                        case "Feature":
                            workItemOffsets.Epic.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Epic.X;
                            newInsertionPoint.Y = workItemOffsets.Epic.Y;
                            break;

                        case "Release":
                            workItemOffsets.Epic.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Epic.X;
                            newInsertionPoint.Y = workItemOffsets.Epic.Y;
                            break;

                        case "Requirement":
                            workItemOffsets.Epic.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Epic.X;
                            newInsertionPoint.Y = workItemOffsets.Epic.Y;
                            break;

                        case "Task":
                            workItemOffsets.Epic.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Epic.X;
                            newInsertionPoint.Y = workItemOffsets.Epic.Y;
                            break;

                        case "Test Case":
                            workItemOffsets.Epic.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Epic.X;
                            newInsertionPoint.Y = workItemOffsets.Epic.Y;
                            break;

                        case "User Needs":
                            workItemOffsets.Epic.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Epic.X;
                            newInsertionPoint.Y = workItemOffsets.Epic.Y;
                            break;

                        case "User Story":
                            workItemOffsets.Epic.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Epic.X;
                            newInsertionPoint.Y = workItemOffsets.Epic.Y;
                            break;

                        default:
                            break;
                    }

                    break;

                case "Feature":
                    switch (shapeWorkItemType)
                    {
                        case "Bug":
                            workItemOffsets.Feature.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Feature.X;
                            newInsertionPoint.Y = workItemOffsets.Feature.Y;
                            break;

                        case "Epic":
                            workItemOffsets.Feature.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Feature.X;
                            newInsertionPoint.Y = workItemOffsets.Feature.Y;
                            break;

                        case "Feature":

                            workItemOffsets.Feature.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Feature.X;
                            newInsertionPoint.Y = workItemOffsets.Feature.Y;
                            break;

                        case "Release":
                            //workItemOffsets.Feature.DecrementHorizontal(width);
                            //newInsertionPoint.X = workItemOffsets.Feature.X;
                            //newInsertionPoint.Y = workItemOffsets.Feature.Y;
                            if (workItemOffsets.UserNeeds.Count > 0)
                            {
                                workItemOffsets.UserStory.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.UserNeeds.X;
                                newInsertionPoint.Y = workItemOffsets.UserNeeds.Y;
                            }
                            else if (workItemOffsets.UserStory.Count > 0)
                            {
                                workItemOffsets.UserStory.IncrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }
                            else
                            {
                                workItemOffsets.Feature.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Feature.X;
                                newInsertionPoint.Y = workItemOffsets.Feature.Y;
                            }
                            break;

                        case "Requirement":
                            workItemOffsets.Feature.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Feature.X;
                            newInsertionPoint.Y = workItemOffsets.Feature.Y;
                            break;

                        case "Task":
                            workItemOffsets.Feature.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Feature.X;
                            newInsertionPoint.Y = workItemOffsets.Feature.Y;
                            break;

                        case "Test Case":
                            workItemOffsets.Feature.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Feature.X;
                            newInsertionPoint.Y = workItemOffsets.Feature.Y;
                            break;

                        case "User Needs":
                            workItemOffsets.Feature.DecrementHorizontal(width, OffsetDirection.Up);
                            newInsertionPoint.X = workItemOffsets.Feature.X;
                            newInsertionPoint.Y = workItemOffsets.Feature.Y;
                            break;

                        case "User Story":
                            workItemOffsets.Feature.DecrementHorizontal(width, OffsetDirection.Up);
                            newInsertionPoint.X = workItemOffsets.Feature.X;
                            newInsertionPoint.Y = workItemOffsets.Feature.Y;
                            break;

                        default:
                            break;
                    }

                    break;

                case "Release":
                    switch (shapeWorkItemType)
                    {
                        case "Bug":
                            if (workItemOffsets.UserStory.Count > 0)
                            {
                                workItemOffsets.UserStory.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }
                            else
                            {
                                workItemOffsets.Release.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Release.X;
                                newInsertionPoint.Y = workItemOffsets.Release.Y;
                            }

                            break;

                        case "Epic":
                            workItemOffsets.Release.IncrementHorizontal(width);

                            newInsertionPoint.X = workItemOffsets.Release.X;
                            newInsertionPoint.Y = workItemOffsets.Release.Y;
                            break;

                        case "Feature":
                            if (workItemOffsets.UserNeeds.Count > 0)
                            {
                                workItemOffsets.UserNeeds.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.UserNeeds.X;
                                newInsertionPoint.Y = workItemOffsets.UserNeeds.Y;
                            }
                            else
                            {
                                workItemOffsets.Release.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Release.X;
                                newInsertionPoint.Y = workItemOffsets.Release.Y;
                            }
                            break;

                        case "Release":
                            workItemOffsets.Release.DecrementHorizontal(width, OffsetDirection.Up);

                            newInsertionPoint.X = workItemOffsets.Release.X;
                            newInsertionPoint.Y = workItemOffsets.Release.Y;
                            break;

                        case "Requirement":
                            if (workItemOffsets.UserNeeds.Count > 0)
                            {
                                workItemOffsets.UserNeeds.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.UserNeeds.X;
                                newInsertionPoint.Y = workItemOffsets.UserNeeds.Y;
                            }
                            else
                            {
                                workItemOffsets.Release.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Release.X;
                                newInsertionPoint.Y = workItemOffsets.Release.Y;
                            }
                            break;

                        case "Task":
                            workItemOffsets.Release.IncrementHorizontal(width, OffsetDirection.Down);

                            newInsertionPoint.X = workItemOffsets.Release.X;
                            newInsertionPoint.Y = workItemOffsets.Release.Y;
                            break;

                        case "Test Case":
                            if (workItemOffsets.UserStory.Count > 0)
                            {
                                workItemOffsets.UserStory.DecrementHorizontal(width, OffsetDirection.Left);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }
                            else
                            {
                                workItemOffsets.Release.DecrementHorizontal(width, OffsetDirection.Left);
                                newInsertionPoint.X = workItemOffsets.Release.X;
                                newInsertionPoint.Y = workItemOffsets.Release.Y;
                            }
                            break;

                        case "Test Plan":
                            workItemOffsets.Release.DecrementHorizontal(width, OffsetDirection.Left);
                            newInsertionPoint.X = workItemOffsets.Release.X;
                            newInsertionPoint.Y = workItemOffsets.Release.Y;
                            break;

                        case "Test Suite":
                            workItemOffsets.Release.DecrementHorizontal(width, OffsetDirection.Left);
                            newInsertionPoint.X = workItemOffsets.Release.X;
                            newInsertionPoint.Y = workItemOffsets.Release.Y;
                            break;

                        case "User Needs":
                            workItemOffsets.Release.DecrementHorizontal(width, OffsetDirection.Up);
                            newInsertionPoint.X = workItemOffsets.Release.X;
                            newInsertionPoint.Y = workItemOffsets.Release.Y;
                            break;

                        case "User Story":
                            if (workItemOffsets.Feature.Count > 0)
                            {
                                workItemOffsets.Feature.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Feature.X;
                                newInsertionPoint.Y = workItemOffsets.Feature.Y;
                            }
                            else
                            {
                                workItemOffsets.Release.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Release.X;
                                newInsertionPoint.Y = workItemOffsets.Release.Y;
                            }

                            break;

                        default:
                            break;
                    }

                    break;

                case "Request":
                    switch (shapeWorkItemType)
                    {
                        case "Bug":
                            workItemOffsets.Request.DecrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Request.X;
                            newInsertionPoint.Y = workItemOffsets.Request.Y;
                            break;

                        case "Epic":
                            workItemOffsets.Request.DecrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Request.X;
                            newInsertionPoint.Y = workItemOffsets.Request.Y;
                            break;

                        case "Feature":
                            workItemOffsets.Request.DecrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Request.X;
                            newInsertionPoint.Y = workItemOffsets.Request.Y;
                            break;

                        case "Release":
                            if (workItemOffsets.UserStory.Count > 0)
                            {
                                workItemOffsets.UserStory.IncrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }
                            else
                            {
                                workItemOffsets.Request.DecrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.Request.X;
                                newInsertionPoint.Y = workItemOffsets.Request.Y;
                            }

                            break;

                        case "Requirement":
                            workItemOffsets.Request.DecrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Request.X;
                            newInsertionPoint.Y = workItemOffsets.Request.Y;
                            break;

                        case "Task":
                            if (workItemOffsets.UserStory.Count > 0)
                            {
                                workItemOffsets.UserStory.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }
                            else if (workItemOffsets.Requirement.Count > 0)
                            {
                                workItemOffsets.Requirement.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Requirement.X;
                                newInsertionPoint.Y = workItemOffsets.Requirement.Y;
                            }
                            else
                            {
                                workItemOffsets.Request.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Request.X;
                                newInsertionPoint.Y = workItemOffsets.Request.Y;
                            }

                            break;

                        case "Test Case":
                            workItemOffsets.Request.DecrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Request.X;
                            newInsertionPoint.Y = workItemOffsets.Request.Y;
                            break;

                        case "User Needs":
                            workItemOffsets.Request.IncrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Request.X;
                            newInsertionPoint.Y = workItemOffsets.Request.Y;
                            break;

                        case "User Story":
                            workItemOffsets.Request.IncrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Request.X;
                            newInsertionPoint.Y = workItemOffsets.Request.Y;
                            break;

                        default:
                            break;
                    }

                    break;

                case "Requirement":
                    switch (shapeWorkItemType)
                    {
                        case "Bug":
                            workItemOffsets.Requirement.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Requirement.X;
                            newInsertionPoint.Y = workItemOffsets.Requirement.Y;
                            break;

                        case "Epic":
                            workItemOffsets.Requirement.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Requirement.X;
                            newInsertionPoint.Y = workItemOffsets.Requirement.Y;
                            break;

                        case "Feature":
                            workItemOffsets.Requirement.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Requirement.X;
                            newInsertionPoint.Y = workItemOffsets.Requirement.Y;
                            break;

                        case "Release":
                            if (workItemOffsets.UserStory.Count > 0)
                            {
                                workItemOffsets.UserStory.IncrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }
                            else
                            {
                                workItemOffsets.Requirement.IncrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.Requirement.X;
                                newInsertionPoint.Y = workItemOffsets.Requirement.Y;
                            }

                            break;

                        case "Requirement":
                            workItemOffsets.Requirement.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Requirement.X;
                            newInsertionPoint.Y = workItemOffsets.Requirement.Y;
                            break;

                        case "Task":
                            workItemOffsets.Requirement.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Requirement.X;
                            newInsertionPoint.Y = workItemOffsets.Requirement.Y;
                            break;

                        case "Test Case":
                            if (workItemOffsets.UserStory.Count > 0)
                            {
                                workItemOffsets.UserStory.DecrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }
                            else
                            {
                                workItemOffsets.Requirement.DecrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.Requirement.X;
                                newInsertionPoint.Y = workItemOffsets.Requirement.Y;
                            }
                            break;

                        case "User Needs":
                            workItemOffsets.Requirement.IncrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Requirement.X;
                            newInsertionPoint.Y = workItemOffsets.Requirement.Y;
                            break;

                        case "User Story":
                            workItemOffsets.Requirement.IncrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Requirement.X;
                            newInsertionPoint.Y = workItemOffsets.Requirement.Y;
                            break;

                        default:
                            break;
                    }

                    break;

                case "Task":
                    switch (shapeWorkItemType)
                    {
                        case "Bug":
                            workItemOffsets.Task.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Task.X;
                            newInsertionPoint.Y = workItemOffsets.Task.Y;
                            break;

                        case "Epic":
                            workItemOffsets.Task.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Task.X;
                            newInsertionPoint.Y = workItemOffsets.Task.Y;
                            break;

                        case "Feature":
                            workItemOffsets.Task.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Task.X;
                            newInsertionPoint.Y = workItemOffsets.Task.Y;
                            break;

                        case "Release":
                            if (workItemOffsets.UserStory.Count > 0)
                            {
                                workItemOffsets.UserStory.IncrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }
                            else
                            {
                                workItemOffsets.Task.IncrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.Task.X;
                                newInsertionPoint.Y = workItemOffsets.Task.Y;
                            }

                            break;

                        case "Request":
                            workItemOffsets.Task.IncrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Task.X;
                            newInsertionPoint.Y = workItemOffsets.Task.Y;
                            break;

                        case "Requirement":
                            workItemOffsets.Task.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Task.X;
                            newInsertionPoint.Y = workItemOffsets.Task.Y;
                            break;

                        case "Task":
                            workItemOffsets.Task.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Task.X;
                            newInsertionPoint.Y = workItemOffsets.Task.Y;
                            break;

                        case "Test Case":
                            workItemOffsets.Task.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.Task.X;
                            newInsertionPoint.Y = workItemOffsets.Task.Y;
                            break;

                        case "User Needs":
                            workItemOffsets.Task.IncrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Task.X;
                            newInsertionPoint.Y = workItemOffsets.Task.Y;
                            break;

                        case "User Story":
                            workItemOffsets.Task.IncrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.Task.X;
                            newInsertionPoint.Y = workItemOffsets.Task.Y;
                            break;

                        default:
                            break;
                    }

                    break;

                case "Test Case":
                    switch (shapeWorkItemType)
                    {
                        case "Bug":
                            if (workItemOffsets.Release.Count > 0)
                            {
                                workItemOffsets.Release.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Release.X;
                                newInsertionPoint.Y = workItemOffsets.Release.Y;
                            }
                            else if (workItemOffsets.UserStory.Count > 0)
                            {
                                workItemOffsets.UserStory.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }
                            else if (workItemOffsets.TestCase.Count > 0)
                            {
                                workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.TestCase.X;
                                newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            }
                            else
                            {
                                workItemOffsets.Unknown.IncrementHorizontal(width);
                                newInsertionPoint.X = workItemOffsets.Unknown.X;
                                newInsertionPoint.Y = workItemOffsets.Unknown.Y;
                            }
                            //workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                            break;

                        case "Epic":
                            workItemOffsets.TestCase.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "Feature":
                            workItemOffsets.TestCase.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "Release":
                            if (workItemOffsets.TestPlan.Count > 0)
                            {
                                workItemOffsets.TestPlan.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.TestPlan.X;
                                newInsertionPoint.Y = workItemOffsets.TestPlan.Y;
                            }
                            else if (workItemOffsets.TestSuite.Count > 0)
                            {
                                workItemOffsets.TestSuite.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.TestSuite.X;
                                newInsertionPoint.Y = workItemOffsets.TestSuite.Y;
                            }
                            else
                            {
                                workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.TestCase.X;
                                newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            }

                            break;

                        case "Requirement":
                            workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "Task":
                            workItemOffsets.TestCase.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "Test Case":
                            workItemOffsets.TestCase.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "User Needs":
                            workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "User Story":
                            workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        default:
                            break;
                    }

                    break;

                case "Test Plan":
                    switch (shapeWorkItemType)
                    {
                        case "Bug":
                            if (workItemOffsets.Release.Count > 0)
                            {
                                workItemOffsets.Release.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Release.X;
                                newInsertionPoint.Y = workItemOffsets.Release.Y;
                            }
                            else if (workItemOffsets.UserStory.Count > 0)
                            {
                                workItemOffsets.UserStory.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }
                            else if (workItemOffsets.TestCase.Count > 0)
                            {
                                workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.TestCase.X;
                                newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            }
                            else
                            {
                                workItemOffsets.Unknown.IncrementHorizontal(width);
                                newInsertionPoint.X = workItemOffsets.Unknown.X;
                                newInsertionPoint.Y = workItemOffsets.Unknown.Y;
                            }
                            //workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                            break;

                        case "Epic":
                            workItemOffsets.TestCase.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "Feature":
                            workItemOffsets.TestCase.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "Release":
                            if (workItemOffsets.TestCase.Count > 0)
                            {
                                workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.TestCase.X;
                                newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            }
                            else if (workItemOffsets.TestSuite.Count > 0)
                            {
                                workItemOffsets.TestSuite.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.TestSuite.X;
                                newInsertionPoint.Y = workItemOffsets.TestSuite.Y;
                            }
                            else
                            {
                                workItemOffsets.TestPlan.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.TestPlan.X;
                                newInsertionPoint.Y = workItemOffsets.TestPlan.Y;
                            }

                            break;

                        case "Requirement":
                            workItemOffsets.TestCase.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "Task":
                            workItemOffsets.TestCase.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "Test Case":
                            workItemOffsets.TestCase.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "User Needs":
                            workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "User Story":
                            workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        default:
                            break;
                    }

                    break;

                case "Test Suite":
                    switch (shapeWorkItemType)
                    {
                        case "Bug":
                            if (workItemOffsets.Release.Count > 0)
                            {
                                workItemOffsets.Release.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Release.X;
                                newInsertionPoint.Y = workItemOffsets.Release.Y;
                            }
                            else if (workItemOffsets.UserStory.Count > 0)
                            {
                                workItemOffsets.UserStory.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }
                            else if (workItemOffsets.TestCase.Count > 0)
                            {
                                workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.TestCase.X;
                                newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            }
                            else
                            {
                                workItemOffsets.Unknown.IncrementHorizontal(width);
                                newInsertionPoint.X = workItemOffsets.Unknown.X;
                                newInsertionPoint.Y = workItemOffsets.Unknown.Y;
                            }
                            //workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                            break;

                        case "Epic":
                            workItemOffsets.TestCase.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "Feature":
                            workItemOffsets.TestCase.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "Release":
                            if (workItemOffsets.TestCase.Count > 0)
                            {
                                workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.TestCase.X;
                                newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            }
                            else if (workItemOffsets.TestPlan.Count > 0)
                            {
                                workItemOffsets.TestPlan.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.TestPlan.X;
                                newInsertionPoint.Y = workItemOffsets.TestPlan.Y;
                            }
                            else
                            {
                                workItemOffsets.TestSuite.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.TestSuite.X;
                                newInsertionPoint.Y = workItemOffsets.TestSuite.Y;
                            }

                            break;

                        case "Requirement":
                            workItemOffsets.TestCase.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "Task":
                            workItemOffsets.TestCase.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "Test Case":
                            workItemOffsets.TestCase.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "Test Suite":
                            workItemOffsets.TestSuite.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.TestSuite.X;
                            newInsertionPoint.Y = workItemOffsets.TestSuite.Y;
                            break;

                        case "User Needs":
                            workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        case "User Story":
                            workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                            newInsertionPoint.X = workItemOffsets.TestCase.X;
                            newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            break;

                        default:
                            break;
                    }

                    break;

                case "User Needs":
                    switch (shapeWorkItemType)
                    {
                        case "Bug":
                            workItemOffsets.UserNeeds.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.UserNeeds.X;
                            newInsertionPoint.Y = workItemOffsets.UserNeeds.Y;
                            break;

                        case "UserNeeds":
                            workItemOffsets.UserNeeds.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.UserNeeds.X;
                            newInsertionPoint.Y = workItemOffsets.UserNeeds.Y;
                            break;

                        case "Feature":
                            if (workItemOffsets.Release.Count > 0)
                            {
                                workItemOffsets.Release.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Release.X;
                                newInsertionPoint.Y = workItemOffsets.Release.Y;
                            }
                            else
                            {
                                workItemOffsets.UserNeeds.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.UserNeeds.X;
                                newInsertionPoint.Y = workItemOffsets.UserNeeds.Y;
                            }

                            break;

                        case "Release":
                            if (workItemOffsets.UserStory.Count > 0)
                            {
                                workItemOffsets.UserStory.IncrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }
                            else
                            {
                                workItemOffsets.UserNeeds.IncrementHorizontal(width);
                                newInsertionPoint.X = workItemOffsets.UserNeeds.X;
                                newInsertionPoint.Y = workItemOffsets.UserNeeds.Y;
                            }
                            break;

                        case "Requirement":
                            if (workItemOffsets.Release.Count > 0)
                            {
                                workItemOffsets.Release.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Release.X;
                                newInsertionPoint.Y = workItemOffsets.Release.Y;
                            }
                            else
                            {
                                workItemOffsets.UserNeeds.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.UserNeeds.X;
                                newInsertionPoint.Y = workItemOffsets.UserNeeds.Y;
                            }

                            break;

                        case "Task":
                            workItemOffsets.UserNeeds.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.UserNeeds.X;
                            newInsertionPoint.Y = workItemOffsets.UserNeeds.Y;
                            break;

                        case "Test Case":
                            workItemOffsets.UserNeeds.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.UserNeeds.X;
                            newInsertionPoint.Y = workItemOffsets.UserNeeds.Y;
                            break;

                        case "User Needs":
                            workItemOffsets.UserNeeds.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.UserNeeds.X;
                            newInsertionPoint.Y = workItemOffsets.UserNeeds.Y;
                            break;

                        case "User Story":
                            workItemOffsets.UserNeeds.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.UserNeeds.X;
                            newInsertionPoint.Y = workItemOffsets.UserNeeds.Y;
                            break;

                        default:
                            break;
                    }

                    break;
                case "User Story":
                    switch (shapeWorkItemType)
                    {
                        case "Bug":
                            if (workItemOffsets.TestCase.Count > 0)
                            {
                                workItemOffsets.TestCase.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.TestCase.X;
                                newInsertionPoint.Y = workItemOffsets.TestCase.Y;
                            }
                            else if (workItemOffsets.UserStory.Count > 0)
                            {
                                workItemOffsets.UserStory.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }
                            else if (workItemOffsets.Release.Count > 0)
                            {
                                workItemOffsets.Release.IncrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Release.X;
                                newInsertionPoint.Y = workItemOffsets.Release.Y;
                            }
                            else
                            {
                                workItemOffsets.Unknown.IncrementHorizontal(width);
                                newInsertionPoint.X = workItemOffsets.Unknown.X;
                                newInsertionPoint.Y = workItemOffsets.Unknown.Y;
                            }

                            break;

                        case "Epic":
                            workItemOffsets.UserStory.IncrementHorizontal(width);
                            break;

                        case "Feature":
                            workItemOffsets.UserStory.IncrementHorizontal(width, OffsetDirection.Down);
                            newInsertionPoint.X = workItemOffsets.UserStory.X;
                            newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            break;

                        case "Release":
                            if (workItemOffsets.Requirement.Count > 0)
                            {
                                workItemOffsets.Requirement.IncrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.Requirement.X;
                                newInsertionPoint.Y = workItemOffsets.Requirement.Y;
                            }
                            else
                            {
                                workItemOffsets.UserStory.IncrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }

                            break;

                        case "Request":
                            workItemOffsets.UserStory.DecrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.UserStory.X;
                            newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            break;

                        case "Requirement":
                            workItemOffsets.UserStory.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.UserStory.X;
                            newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            break;

                        case "Task":
                            if (workItemOffsets.Requirement.Count > 0)
                            {
                                workItemOffsets.Requirement.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Requirement.X;
                                newInsertionPoint.Y = workItemOffsets.Requirement.Y;
                            }
                            else if (workItemOffsets.Request.Count > 0)
                            {
                                workItemOffsets.Request.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Request.X;
                                newInsertionPoint.Y = workItemOffsets.Request.Y;
                            }
                            else if (workItemOffsets.ProductionIssue.Count > 0)
                            {
                                workItemOffsets.ProductionIssue.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.ProductionIssue.X;
                                newInsertionPoint.Y = workItemOffsets.ProductionIssue.Y;
                            }
                            else if (workItemOffsets.Issue.Count > 0)
                            {
                                workItemOffsets.Issue.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.Issue.X;
                                newInsertionPoint.Y = workItemOffsets.Issue.Y;
                            }
                            else
                            {
                                workItemOffsets.UserStory.DecrementHorizontal(width, OffsetDirection.Up);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }

                            break;

                        case "Test Case":
                            if (workItemOffsets.Bug.Count > 0)
                            {
                                workItemOffsets.Bug.DecrementHorizontal(width, OffsetDirection.Down);
                            }
                            else
                            {
                                workItemOffsets.UserStory.DecrementHorizontal(width, OffsetDirection.Down);
                                newInsertionPoint.X = workItemOffsets.UserStory.X;
                                newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            }
                            break;

                        case "User Needs":
                            workItemOffsets.UserStory.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.UserStory.X;
                            newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            break;

                        case "User Story":
                            workItemOffsets.UserStory.IncrementHorizontal(width);
                            newInsertionPoint.X = workItemOffsets.UserStory.X;
                            newInsertionPoint.Y = workItemOffsets.UserStory.Y;
                            break;

                        default:
                            break;
                    }

                    break;

                default:
                    newInsertionPoint.X = initialPosition.X;
                    newInsertionPoint.Y = initialPosition.Y;
                    break;
            }

            return newInsertionPoint;
        }

        internal static Point CalculateInsertionPointQueriedWorkItems(Point initialPosition, 
            WorkItem linkedWorkItem, WorkItemShapeInfo activeShape, WorkItemOffsets workItemOffsets)
        {
            Point newInsertionPoint = new Point();

            double height = activeShape.Height;
            //double width = activeShape.Width;
            // HACK(crhodes)
            // We need the width of the existing shape.  Hard code for now.

            double width = 0.75;

            string shapeWorkItemType = activeShape.WorkItemType;

            switch (linkedWorkItem.Fields["System.WorkItemType"])
            {
                case "Bug":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Change Request":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Code Review Response":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Code Review Request":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Design Review Request":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Epic":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Feature":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Issue":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Meeting Minutes":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Milestone":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Production Issue":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Release":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Request":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Requirement":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Review":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Review Request":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Specification":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Shared Steps":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Task":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Test Case":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Test Plan":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "Test Suite":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "User Needs":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                case "User Story":
                    workItemOffsets.QueryResult.IncrementHorizontal(width, OffsetDirection.Down, 10);
                    newInsertionPoint.X = workItemOffsets.QueryResult.X;
                    newInsertionPoint.Y = workItemOffsets.QueryResult.Y;

                    break;

                default:
                    workItemOffsets.Unknown.DecrementHorizontal(width, OffsetDirection.Up);
                    newInsertionPoint.X = workItemOffsets.Unknown.X;
                    newInsertionPoint.Y = workItemOffsets.Unknown.Y;
                    break;
            }

            return newInsertionPoint;
        }
    }
}

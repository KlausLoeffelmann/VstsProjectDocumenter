using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VstsProjectDocumenter.DataStructures
{
    public class AdopWorkItemBase
    {
        internal AdopWorkItemBase(int id, 
            Guid identifier, 
            string url, 
            string name)
        {
            Id = id;
            Identifier = identifier;
            Url = url;
            Name = name;
        }

        public int Id { get; }
        public Guid Identifier { get; }
        public string Url { get; }
        public string Name { get; }
    }

    public class AdopIteration : AdopWorkItemBase
    {
        public AdopIteration(WorkItemClassificationNode workItemNode) :
            base(workItemNode.Id, workItemNode.Identifier, workItemNode.Url, workItemNode.Name)
        {
            StartDate = (DateTime)workItemNode.Attributes["startDate"];
            FinishDate = (DateTime)workItemNode.Attributes["finishDate"];
        }

        public DateTime StartDate { get; }
        public DateTime FinishDate { get; }
        public int Level {get; internal set;}

        public static ImmutableList<AdopIteration> GetIterations(IEnumerable<WorkItemClassificationNode> nodes)
        {
            var list = ImmutableList.CreateBuilder<AdopIteration>();

            // Simplification: We're using only and always 3 Levels , so no recursion.
            // And we're need a flattened list.
            // TODO: This could/should be done recursively.
            foreach (var nodeItem in nodes)
            {
                try
                {
                    foreach (var childNodeItem in nodeItem.Children)
                    {
                        list.Add(new AdopIteration(childNodeItem) { Level = 1 });
                        var children = childNodeItem?.Children;
                        if ((children is object))
                        {
                            foreach (var childChildNodeItem in children)
                            {
                                list.Add(new AdopIteration(childChildNodeItem) { Level = 2 });
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                }
            }

            return list.ToImmutable();
        }
    }
}

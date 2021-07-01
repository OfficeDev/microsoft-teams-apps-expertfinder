// <copyright file="SearchPropertiesResult.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Models.SharePoint
{
    /// <summary>
    /// Properties result data from SharePoint search response cell data model.
    /// </summary>
    public class SearchPropertiesResult
    {
        /// <summary>
        /// Gets or sets key value.
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// Gets or sets value.
        /// </summary>
        public string Value { get; set; }

        /// <summary>
        /// Gets or sets data type of value.
        /// </summary>
        public string ValueType { get; set; }
    }
}

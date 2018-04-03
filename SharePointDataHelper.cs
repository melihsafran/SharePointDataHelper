using System;
using System.Collections.Generic;
using Microsoft.SharePoint;

public static class SharePointDataHelper
{
    // SharePointDataHelper
    // Melih SAFRAN
    // Create: 09.12.2014
    // Update: 09.03.2016

    public static SPListItemCollection GetItems(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listName, SPQuery spQuery)
    {
        SPListItemCollection spListItemCollectionResult = null;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.Lists[listName];

                            spListItemCollectionResult = spQuery == null ? spList.GetItems() : spList.GetItems(spQuery);
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.Lists[listName];

                        spListItemCollectionResult = spQuery == null ? spList.GetItems() : spList.GetItems(spQuery);
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return spListItemCollectionResult;
    }

    public static SPListItemCollection GetItemsByInternalListName(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listInternalName, SPQuery spQuery)
    {
        SPListItemCollection spListItemCollectionResult = null;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);

                            spListItemCollectionResult = spQuery == null ? spList.GetItems() : spList.GetItems(spQuery);
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);

                        spListItemCollectionResult = spQuery == null ? spList.GetItems() : spList.GetItems(spQuery);
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return spListItemCollectionResult;
    }


    public static int GetItemsCount(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listName, SPQuery spQuery)
    {
        int spListItemCollectionResultCount = 0;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.Lists[listName];

                            spListItemCollectionResultCount = spQuery == null ? spList.ItemCount : spList.GetItems(spQuery).Count;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.Lists[listName];

                        spListItemCollectionResultCount = spQuery == null ? spList.ItemCount : spList.GetItems(spQuery).Count;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return spListItemCollectionResultCount;
    }


    public static int GetItemsCountByInternalListName(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listInternalName, SPQuery spQuery)
    {
        int spListItemCollectionResultCount = 0;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);

                            spListItemCollectionResultCount = spQuery == null ? spList.ItemCount : spList.GetItems(spQuery).Count;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);

                        spListItemCollectionResultCount = spQuery == null ? spList.ItemCount : spList.GetItems(spQuery).Count;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return spListItemCollectionResultCount;
    }


    public static SPListItem GetItemByID(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listName, int id)
    {
        SPListItem spListItemResult = null;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.Lists[listName];

                            spListItemResult = spList.GetItemById(id);
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.Lists[listName];

                        spListItemResult = spList.GetItemById(id);
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return spListItemResult;
    }


    public static SPListItem GetItemByIDByInternalListName(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listInternalName, int id)
    {
        SPListItem spListItemResult = null;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);

                            spListItemResult = spList.GetItemById(id);
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);

                        spListItemResult = spList.GetItemById(id);
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return spListItemResult;
    }


    public static bool AddItem(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listName, Dictionary<string, object> parameters)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.Lists[listName];

                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;
                            SPListItem spListItem = spList.AddItem();

                            foreach (var parameter in parameters)
                            {
                                spListItem[parameter.Key] = parameter.Value;
                            }

                            spListItem.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.Lists[listName];
                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;
                        spWeb.AllowUnsafeUpdates = true;
                        SPListItem spListItem = spList.AddItem();

                        foreach (var parameter in parameters)
                        {
                            spListItem[parameter.Key] = parameter.Value;
                        }

                        spListItem.Update();
                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }


    public static bool AddItemByInternalListName(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listInternalName, Dictionary<string, object> parameters)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);

                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;
                            SPListItem spListItem = spList.AddItem();

                            foreach (var parameter in parameters)
                            {
                                spListItem[parameter.Key] = parameter.Value;
                            }

                            spListItem.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);
                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;
                        spWeb.AllowUnsafeUpdates = true;
                        SPListItem spListItem = spList.AddItem();

                        foreach (var parameter in parameters)
                        {
                            spListItem[parameter.Key] = parameter.Value;
                        }

                        spListItem.Update();
                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static int AddItemReturnID(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listName, Dictionary<string, object> parameters)
    {
        int result = 0;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.Lists[listName];
                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;
                            spWeb.AllowUnsafeUpdates = true;
                            SPListItem spListItem = spList.AddItem();

                            foreach (var parameter in parameters)
                            {
                                spListItem[parameter.Key] = parameter.Value;
                            }

                            spListItem.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = spListItem.ID;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.Lists[listName];
                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                        spWeb.AllowUnsafeUpdates = true;
                        SPListItem spListItem = spList.AddItem();

                        foreach (var parameter in parameters)
                        {
                            spListItem[parameter.Key] = parameter.Value;
                        }

                        spListItem.Update();
                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = spListItem.ID;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static int AddItemReturnIDByInternalListName(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listInternalName, Dictionary<string, object> parameters)
    {
        int result = 0;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);
                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;
                            spWeb.AllowUnsafeUpdates = true;
                            SPListItem spListItem = spList.AddItem();

                            foreach (var parameter in parameters)
                            {
                                spListItem[parameter.Key] = parameter.Value;
                            }

                            spListItem.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = spListItem.ID;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);
                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                        spWeb.AllowUnsafeUpdates = true;
                        SPListItem spListItem = spList.AddItem();

                        foreach (var parameter in parameters)
                        {
                            spListItem[parameter.Key] = parameter.Value;
                        }

                        spListItem.Update();
                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = spListItem.ID;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool UpdateItem(bool runWithElevatedPrivileges, SPListItem spListItem, Dictionary<string, object> parameters)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(spListItem.Web.Site.Url))
                    {
                        using (SPWeb spWeb = spSite.OpenWeb(spListItem.Web.Url))
                        {
                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;

                            foreach (var parameter in parameters)
                            {
                                spListItem[parameter.Key] = parameter.Value;
                            }

                            spListItem.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(spListItem.Web.Site.Url))
                {
                    using (SPWeb spWeb = spSite.OpenWeb(spListItem.Web.Url))
                    {
                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                        spWeb.AllowUnsafeUpdates = true;

                        foreach (var parameter in parameters)
                        {
                            spListItem[parameter.Key] = parameter.Value;
                        }

                        spListItem.Update();
                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool UpdateItemByID(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listName, int spListItemID, Dictionary<string, object> parameters)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPListItem spListItem = spWeb.Lists[listName].GetItemById(spListItemID);

                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;

                            foreach (var parameter in parameters)
                            {
                                spListItem[parameter.Key] = parameter.Value;
                            }

                            spListItem.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPListItem spListItem = spWeb.Lists[listName].GetItemById(spListItemID);

                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                        spWeb.AllowUnsafeUpdates = true;

                        foreach (var parameter in parameters)
                        {
                            spListItem[parameter.Key] = parameter.Value;
                        }

                        spListItem.Update();
                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool UpdateItemByIDByInternalListName(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listInternalName, int spListItemID, Dictionary<string, object> parameters)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPListItem spListItem = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName).GetItemById(spListItemID);

                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;

                            foreach (var parameter in parameters)
                            {
                                spListItem[parameter.Key] = parameter.Value;
                            }

                            spListItem.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPListItem spListItem = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName).GetItemById(spListItemID);

                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                        spWeb.AllowUnsafeUpdates = true;

                        foreach (var parameter in parameters)
                        {
                            spListItem[parameter.Key] = parameter.Value;
                        }

                        spListItem.Update();
                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool UpdateItemBySPQuery(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listName, SPQuery spQuery, Dictionary<string, object> parameters)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPListItemCollection spListItemCollection = spWeb.Lists[listName].GetItems(spQuery);

                            if (spListItemCollection.Count == 1)
                            {
                                SPListItem spListItem = spListItemCollection[0];

                                bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                                spWeb.AllowUnsafeUpdates = true;

                                foreach (var parameter in parameters)
                                {
                                    spListItem[parameter.Key] = parameter.Value;
                                }

                                spListItem.Update();
                                spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                                result = true;
                            }
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPListItemCollection spListItemCollection = spWeb.Lists[listName].GetItems(spQuery);

                        if (spListItemCollection.Count == 1)
                        {
                            SPListItem spListItem = spListItemCollection[0];

                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;

                            foreach (var parameter in parameters)
                            {
                                spListItem[parameter.Key] = parameter.Value;
                            }

                            spListItem.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool UpdateItemBySPQueryByInternalListName(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listInternalName, SPQuery spQuery, Dictionary<string, object> parameters)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPListItemCollection spListItemCollection = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName).GetItems(spQuery);

                            if (spListItemCollection.Count == 1)
                            {
                                SPListItem spListItem = spListItemCollection[0];

                                bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                                spWeb.AllowUnsafeUpdates = true;

                                foreach (var parameter in parameters)
                                {
                                    spListItem[parameter.Key] = parameter.Value;
                                }

                                spListItem.Update();
                                spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                                result = true;
                            }
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPListItemCollection spListItemCollection = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName).GetItems(spQuery);

                        if (spListItemCollection.Count == 1)
                        {
                            SPListItem spListItem = spListItemCollection[0];

                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;

                            foreach (var parameter in parameters)
                            {
                                spListItem[parameter.Key] = parameter.Value;
                            }

                            spListItem.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool DeleteItem(bool runWithElevatedPrivileges, SPListItem spListItem)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(spListItem.Web.Site.Url))
                    {
                        using (SPWeb spWeb = spSite.OpenWeb(spListItem.Web.Url))
                        {
                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;

                            spListItem.Delete();

                            spWeb.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(spListItem.Web.Site.Url))
                {
                    using (SPWeb spWeb = spSite.OpenWeb(spListItem.Web.Url))
                    {
                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                        spWeb.AllowUnsafeUpdates = true;

                        spListItem.Delete();

                        spWeb.Update();
                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool DeleteItemByID(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listName, int listItemID)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPListItem spListItem = spWeb.Lists[listName].GetItemById(listItemID);

                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;

                            spListItem.Delete();

                            spWeb.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPListItem spListItem = spWeb.Lists[listName].GetItemById(listItemID);

                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                        spWeb.AllowUnsafeUpdates = true;

                        spListItem.Delete();

                        spWeb.Update();
                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool DeleteItemByIDByInternalListName(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listInternalName, int listItemID)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPListItem spListItem = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName).GetItemById(listItemID);

                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;

                            spListItem.Delete();

                            spWeb.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPListItem spListItem = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName).GetItemById(listItemID);

                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                        spWeb.AllowUnsafeUpdates = true;

                        spListItem.Delete();

                        spWeb.Update();
                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool DeleteItemBySPQuery(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listName, SPQuery spQuery)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPListItemCollection spListItemCollection = spWeb.Lists[listName].GetItems(spQuery);

                            if (spListItemCollection.Count > 0)
                            {
                                bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                                spWeb.AllowUnsafeUpdates = true;

                                List<SPListItem> listSPListItem = new List<SPListItem>();
                                foreach (SPListItem spListItem in spListItemCollection)
                                {
                                    listSPListItem.Add(spListItem);
                                }

                                foreach (SPListItem spListItem in listSPListItem)
                                {
                                    spListItem.Delete();
                                }

                                spWeb.Update();
                                spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                                result = true;
                            }
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPListItemCollection spListItemCollection = spWeb.Lists[listName].GetItems(spQuery);

                        if (spListItemCollection.Count > 0)
                        {
                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;

                            List<SPListItem> listSPListItem = new List<SPListItem>();
                            foreach (SPListItem spListItem in spListItemCollection)
                            {
                                listSPListItem.Add(spListItem);
                            }

                            foreach (SPListItem spListItem in listSPListItem)
                            {
                                spListItem.Delete();
                            }

                            spWeb.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool DeleteItemBySPQueryByInternalListName(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listInternalName, SPQuery spQuery)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPListItemCollection spListItemCollection = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName).GetItems(spQuery);

                            if (spListItemCollection.Count > 0)
                            {
                                bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                                spWeb.AllowUnsafeUpdates = true;

                                List<SPListItem> listSPListItem = new List<SPListItem>();
                                foreach (SPListItem spListItem in spListItemCollection)
                                {
                                    listSPListItem.Add(spListItem);
                                }

                                foreach (SPListItem spListItem in listSPListItem)
                                {
                                    spListItem.Delete();
                                }

                                spWeb.Update();
                                spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                                result = true;
                            }
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPListItemCollection spListItemCollection = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName).GetItems(spQuery);

                        if (spListItemCollection.Count > 0)
                        {
                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;

                            List<SPListItem> listSPListItem = new List<SPListItem>();
                            foreach (SPListItem spListItem in spListItemCollection)
                            {
                                listSPListItem.Add(spListItem);
                            }

                            foreach (SPListItem spListItem in listSPListItem)
                            {
                                spListItem.Delete();
                            }

                            spWeb.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool DeleteAllItems(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listName)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.Lists[listName];
                            SPListItemCollection spListItemCollection = spList.GetItems();

                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;

                            List<SPListItem> listSPListItem = new List<SPListItem>();
                            foreach (SPListItem spListItem in spListItemCollection)
                            {
                                listSPListItem.Add(spListItem);
                            }

                            foreach (SPListItem spListItem in listSPListItem)
                            {
                                spListItem.Delete();
                            }

                            spWeb.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.Lists[listName];
                        SPListItemCollection spListItemCollection = spList.GetItems();

                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                        spWeb.AllowUnsafeUpdates = true;

                        List<SPListItem> listSPListItem = new List<SPListItem>();
                        foreach (SPListItem spListItem in spListItemCollection)
                        {
                            listSPListItem.Add(spListItem);
                        }

                        foreach (SPListItem spListItem in listSPListItem)
                        {
                            spListItem.Delete();
                        }

                        spWeb.Update();
                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool DeleteAllItemsByInternalListName(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listInternalName)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);
                            SPListItemCollection spListItemCollection = spList.GetItems();

                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;

                            List<SPListItem> listSPListItem = new List<SPListItem>();
                            foreach (SPListItem spListItem in spListItemCollection)
                            {
                                listSPListItem.Add(spListItem);
                            }

                            foreach (SPListItem spListItem in listSPListItem)
                            {
                                spListItem.Delete();
                            }

                            spWeb.Update();
                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);
                        SPListItemCollection spListItemCollection = spList.GetItems();

                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                        spWeb.AllowUnsafeUpdates = true;

                        List<SPListItem> listSPListItem = new List<SPListItem>();
                        foreach (SPListItem spListItem in spListItemCollection)
                        {
                            listSPListItem.Add(spListItem);
                        }

                        foreach (SPListItem spListItem in listSPListItem)
                        {
                            spListItem.Delete();
                        }

                        spWeb.Update();
                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool DeleteBulkAllItems(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listName, SPQuery spQuery, uint deleteBatchCount)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.Lists[listName];

                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;

                            string command = String.Format(@"<Method><SetList>{0}</SetList><SetVar Name=""ID"">{1}</SetVar><SetVar Name=""Cmd"">Delete</SetVar></Method>", spList.ID, "{0}");

                            if (spQuery == null)
                            {
                                spQuery = new SPQuery();
                                spQuery.RowLimit = deleteBatchCount;
                            }

                            while (spList.ItemCount > 0)
                            {
                                SPListItemCollection spListItemCollection = spList.GetItems(spQuery);

                                System.Text.StringBuilder stringBuilderDelete = new System.Text.StringBuilder();
                                stringBuilderDelete.Append(@"<?xml version=""1.0"" encoding=""UTF-8""?><Batch>");

                                Guid[] deletedListItemIDs = new Guid[spListItemCollection.Count];

                                for (int i = 0; i < spListItemCollection.Count; i++)
                                {
                                    SPListItem spListItem = spListItemCollection[i];
                                    stringBuilderDelete.Append(string.Format(command, spListItem.ID.ToString()));
                                    deletedListItemIDs[i] = spListItem.UniqueId;
                                }
                                stringBuilderDelete.Append("</Batch>");

                                spWeb.ProcessBatchData(stringBuilderDelete.ToString());

                                //if (deleteFromRecycleBin)
                                //{
                                //    spWeb.RecycleBin.Delete(deletedListItemIDs);
                                //}

                                spList.Update();
                            }

                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.Lists[listName];

                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                        spWeb.AllowUnsafeUpdates = true;

                        string command = String.Format(@"<Method><SetList>{0}</SetList><SetVar Name=""ID"">{1}</SetVar><SetVar Name=""Cmd"">Delete</SetVar></Method>", spList.ID, "{0}");

                        if (spQuery == null)
                        {
                            spQuery = new SPQuery();
                            spQuery.RowLimit = deleteBatchCount;
                        }

                        while (spList.ItemCount > 0)
                        {
                            SPListItemCollection spListItemCollection = spList.GetItems(spQuery);

                            System.Text.StringBuilder stringBuilderDelete = new System.Text.StringBuilder();
                            stringBuilderDelete.Append(@"<?xml version=""1.0"" encoding=""UTF-8""?><Batch>");

                            Guid[] deletedListItemIDs = new Guid[spListItemCollection.Count];

                            for (int i = 0; i < spListItemCollection.Count; i++)
                            {
                                SPListItem spListItem = spListItemCollection[i];
                                stringBuilderDelete.Append(string.Format(command, spListItem.ID.ToString()));
                                deletedListItemIDs[i] = spListItem.UniqueId;
                            }
                            stringBuilderDelete.Append("</Batch>");

                            spWeb.ProcessBatchData(stringBuilderDelete.ToString());

                            //if (deleteFromRecycleBin)
                            //{
                            //    spWeb.RecycleBin.Delete(deletedListItemIDs);
                            //}

                            spList.Update();
                        }

                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool DeleteBulkAllItemsByInternalListName(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listInternalName, SPQuery spQuery, uint deleteBatchCount)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);

                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                            spWeb.AllowUnsafeUpdates = true;

                            string command = String.Format(@"<Method><SetList>{0}</SetList><SetVar Name=""ID"">{1}</SetVar><SetVar Name=""Cmd"">Delete</SetVar></Method>", spList.ID, "{0}");

                            if (spQuery == null)
                            {
                                spQuery = new SPQuery();
                                spQuery.RowLimit = deleteBatchCount;
                            }

                            while (spList.ItemCount > 0)
                            {
                                SPListItemCollection spListItemCollection = spList.GetItems(spQuery);

                                System.Text.StringBuilder stringBuilderDelete = new System.Text.StringBuilder();
                                stringBuilderDelete.Append(@"<?xml version=""1.0"" encoding=""UTF-8""?><Batch>");

                                Guid[] deletedListItemIDs = new Guid[spListItemCollection.Count];

                                for (int i = 0; i < spListItemCollection.Count; i++)
                                {
                                    SPListItem spListItem = spListItemCollection[i];
                                    stringBuilderDelete.Append(string.Format(command, spListItem.ID.ToString()));
                                    deletedListItemIDs[i] = spListItem.UniqueId;
                                }
                                stringBuilderDelete.Append("</Batch>");

                                spWeb.ProcessBatchData(stringBuilderDelete.ToString());

                                //if (deleteFromRecycleBin)
                                //{
                                //    spWeb.RecycleBin.Delete(deletedListItemIDs);
                                //}

                                spList.Update();
                            }

                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);

                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;

                        spWeb.AllowUnsafeUpdates = true;

                        string command = String.Format(@"<Method><SetList>{0}</SetList><SetVar Name=""ID"">{1}</SetVar><SetVar Name=""Cmd"">Delete</SetVar></Method>", spList.ID, "{0}");

                        if (spQuery == null)
                        {
                            spQuery = new SPQuery();
                            spQuery.RowLimit = deleteBatchCount;
                        }

                        while (spList.ItemCount > 0)
                        {
                            SPListItemCollection spListItemCollection = spList.GetItems(spQuery);

                            System.Text.StringBuilder stringBuilderDelete = new System.Text.StringBuilder();
                            stringBuilderDelete.Append(@"<?xml version=""1.0"" encoding=""UTF-8""?><Batch>");

                            Guid[] deletedListItemIDs = new Guid[spListItemCollection.Count];

                            for (int i = 0; i < spListItemCollection.Count; i++)
                            {
                                SPListItem spListItem = spListItemCollection[i];
                                stringBuilderDelete.Append(string.Format(command, spListItem.ID.ToString()));
                                deletedListItemIDs[i] = spListItem.UniqueId;
                            }
                            stringBuilderDelete.Append("</Batch>");

                            spWeb.ProcessBatchData(stringBuilderDelete.ToString());

                            //if (deleteFromRecycleBin)
                            //{
                            //    spWeb.RecycleBin.Delete(deletedListItemIDs);
                            //}

                            spList.Update();
                        }

                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static SPFileCollection GetItemAttachments(SPListItem spListItem)
    {
        SPFileCollection spFileCollection = null;
        try
        {
            if (spListItem.Attachments.Count > 0)
            {
                SPFolder spFolder = spListItem.Web.Folders["Lists"].SubFolders[spListItem.ListItems.List.Title].SubFolders["Attachments"].SubFolders[spListItem.ID.ToString()];

                spFileCollection = spFolder.Files;
            }
        }
        catch
        {
            throw;
        }
        return spFileCollection;
    }

    public static Dictionary<string, string> GetItemAttachmentsAsUrl(SPListItem spListItem)
    {
        Dictionary<string, string> listAttachment = new Dictionary<string, string>();
        try
        {
            if (spListItem.Attachments.Count > 0)
            {
                SPFolder spFolder = spListItem.Web.Folders["Lists"].SubFolders[spListItem.ListItems.List.Title].SubFolders["Attachments"].SubFolders[spListItem.ID.ToString()];

                string fileUrl = string.Empty;

                foreach (SPFile spFile in spFolder.Files)
                {
                    fileUrl = string.Format("{0}/{1}", spListItem.Web.Url, spFile.Url);
                    listAttachment.Add(spFile.Name, fileUrl);
                }
            }
        }
        catch
        {
            throw;
        }
        return listAttachment;
    }

    public static Dictionary<string, string> GetItemAttachmentByFileNameContains(SPListItem spListItem, string searchString)
    {
        Dictionary<string, string> listAttachment = new Dictionary<string, string>();
        try
        {
            if (spListItem.Attachments.Count > 0)
            {
                SPFolder spFolder = spListItem.Web.Folders["Lists"].SubFolders[spListItem.ListItems.List.Title].SubFolders["Attachments"].SubFolders[spListItem.ID.ToString()];

                string fileUrl = string.Empty;

                foreach (SPFile spFile in spFolder.Files)
                {
                    if (spFile.Name.Contains(searchString))
                    {
                        fileUrl = string.Format("{0}/{1}", spListItem.Web.Url, spFile.Url);
                        listAttachment.Add(spFile.Name, fileUrl);
                        break;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return listAttachment;
    }

    public static SPListItemCollection GetDocumentLibraryItems(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listName, SPQuery spQuery)
    {
        SPListItemCollection SPListItemCollectionResult = null;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPDocumentLibrary spDocumentLibrary = spWeb.Lists[listName] as SPDocumentLibrary;

                            SPListItemCollectionResult = spQuery == null ? spDocumentLibrary.GetItems() : spDocumentLibrary.GetItems(spQuery);
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPDocumentLibrary spDocumentLibrary = spWeb.Lists[listName] as SPDocumentLibrary;

                        SPListItemCollectionResult = spQuery == null ? spDocumentLibrary.GetItems() : spDocumentLibrary.GetItems(spQuery);
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return SPListItemCollectionResult;
    }

    public static SPListItemCollection GetDocumentLibraryItemsByInternalListName(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listInternalName, SPQuery spQuery)
    {
        SPListItemCollection SPListItemCollectionResult = null;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPDocumentLibrary spDocumentLibrary = spWeb.GetList(spWeb.Url + "/" + listInternalName) as SPDocumentLibrary;

                            SPListItemCollectionResult = spQuery == null ? spDocumentLibrary.GetItems() : spDocumentLibrary.GetItems(spQuery);
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPDocumentLibrary spDocumentLibrary = spWeb.GetList(spWeb.Url + "/" + listInternalName) as SPDocumentLibrary;

                        SPListItemCollectionResult = spQuery == null ? spDocumentLibrary.GetItems() : spDocumentLibrary.GetItems(spQuery);
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return SPListItemCollectionResult;
    }

    SPFileCollection GetDocumentLibraryFiles(SPListItem spListItem)
    {
        SPFileCollection spFileCollection = null;
        try
        {
            if (spListItem.Folder != null)
            {
                spFileCollection = spListItem.Folder.Files;
            }
        }
        catch
        {
            throw;
        }
        return spFileCollection;
    }

    public static bool AddFileToDocumentLibrary(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listName, string fileName, System.IO.Stream postedFileInputStreamImage)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;
                            spWeb.AllowUnsafeUpdates = true;

                            SPList spList = spWeb.Lists[listName];
                            spList.RootFolder.Files.Add(fileName, postedFileInputStreamImage);

                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;
                        spWeb.AllowUnsafeUpdates = true;

                        SPList spList = spWeb.Lists[listName];
                        spList.RootFolder.Files.Add(fileName, postedFileInputStreamImage);

                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool AddFileToDocumentLibraryByInternalListName(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listInternalName, string fileName, System.IO.Stream postedFileInputStreamImage)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;
                            spWeb.AllowUnsafeUpdates = true;

                            SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);
                            spList.RootFolder.Files.Add(fileName, postedFileInputStreamImage);

                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;
                        spWeb.AllowUnsafeUpdates = true;

                        SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);
                        spList.RootFolder.Files.Add(fileName, postedFileInputStreamImage);

                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool DeleteFileFromDocumentLibrary(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listName, string fileName)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;
                            spWeb.AllowUnsafeUpdates = true;

                            SPList spList = spWeb.Lists[listName];

                            string fileUrl = string.Format("{0}/{1}/{2}", siteUrl + (string.IsNullOrEmpty(webUrl) ? "" : "/" + spWeb), listName, fileName);

                            spList.RootFolder.Files.Delete(fileUrl);

                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;
                        spWeb.AllowUnsafeUpdates = true;

                        SPList spList = spWeb.Lists[listName];
                        string fileUrl = string.Format("{0}/{1}/{2}", siteUrl + (string.IsNullOrEmpty(webUrl) ? "" : "/" + spWeb), listName, fileName);

                        spList.RootFolder.Files.Delete(fileUrl);

                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static bool DeleteFileFromDocumentLibraryByInternalListName(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listInternalName, string fileName)
    {
        bool result = false;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;
                            spWeb.AllowUnsafeUpdates = true;

                            SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);

                            string fileUrl = string.Format("{0}/{1}/{2}", siteUrl + (string.IsNullOrEmpty(webUrl) ? "" : "/" + spWeb), listInternalName, fileName);

                            spList.RootFolder.Files.Delete(fileUrl);

                            spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                            result = true;
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        bool spWebAllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;
                        spWeb.AllowUnsafeUpdates = true;

                        SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);
                        string fileUrl = string.Format("{0}/{1}/{2}", siteUrl + (string.IsNullOrEmpty(webUrl) ? "" : "/" + spWeb), listInternalName, fileName);

                        spList.RootFolder.Files.Delete(fileUrl);

                        spWeb.AllowUnsafeUpdates = spWebAllowUnsafeUpdates;
                        result = true;
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return result;
    }

    public static List<string> GetSPFieldChoiceItems(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listName, string columnName)
    {
        List<string> listSpFieldChoiceValue = null;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.Lists[listName];

                            SPFieldChoice spFieldChoice = (Microsoft.SharePoint.SPFieldChoice)spList.Fields[columnName];

                            if (spFieldChoice != null)
                            {
                                listSpFieldChoiceValue = new List<string>();

                                foreach (string value in spFieldChoice.Choices)
                                {
                                    listSpFieldChoiceValue.Add(value);
                                }
                            }
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.Lists[listName];

                        SPFieldChoice spFieldChoice = (Microsoft.SharePoint.SPFieldChoice)spList.Fields[columnName];

                        if (spFieldChoice != null)
                        {
                            listSpFieldChoiceValue = new List<string>();

                            foreach (string value in spFieldChoice.Choices)
                            {
                                listSpFieldChoiceValue.Add(value);
                            }
                        }
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return listSpFieldChoiceValue;
    }

    public static List<string> GetSPFieldChoiceItemsByInternalListName(bool runWithElevatedPrivileges, string siteUrl, string webUrl, string listInternalName, string columnName)
    {
        List<string> listSpFieldChoiceValue = null;
        try
        {
            if (runWithElevatedPrivileges)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite spSite = new SPSite(siteUrl))
                    {
                        using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                        {
                            SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);

                            SPFieldChoice spFieldChoice = (Microsoft.SharePoint.SPFieldChoice)spList.Fields[columnName];

                            if (spFieldChoice != null)
                            {
                                listSpFieldChoiceValue = new List<string>();

                                foreach (string value in spFieldChoice.Choices)
                                {
                                    listSpFieldChoiceValue.Add(value);
                                }
                            }
                        }
                    }
                });
            }
            else
            {
                using (SPSite spSite = new SPSite(siteUrl))
                {
                    using (SPWeb spWeb = webUrl == string.Empty ? spSite.OpenWeb() : spSite.OpenWeb(webUrl))
                    {
                        SPList spList = spWeb.GetList(spWeb.Url + "/Lists/" + listInternalName);

                        SPFieldChoice spFieldChoice = (Microsoft.SharePoint.SPFieldChoice)spList.Fields[columnName];

                        if (spFieldChoice != null)
                        {
                            listSpFieldChoiceValue = new List<string>();

                            foreach (string value in spFieldChoice.Choices)
                            {
                                listSpFieldChoiceValue.Add(value);
                            }
                        }
                    }
                }
            }
        }
        catch
        {
            throw;
        }
        return listSpFieldChoiceValue;
    }
}


let
    __VAR_SISTEMA_X_USER_LIST__ = 

        let 

            __VAR_ARRAY_CONECCTION_STRINGS__ = 
                {
                    __VAR_DEFAULT_DOMAIN__,
                    __VAR_API_QUERYSTRING__,
                    __VAR_STRING_SITE__,
                    __VAR_IMPLEMENTATION__
                },


            __LIST_REMOVE_METADATA_USERS__ = 
                {
                    "IsEmailAuthenticationGuestUser",
                    "IsShareByEmailGuestUser", 
                    "PrincipalType"
                    "IsHiddenInUI"
                    "IsSiteAdmin",
                    "Expiration", 
                    "LoginName", 
                    "UserId",
                    "Groups",
                    "Alerts",
                    "Email", 
                    "Title"
                },


            __FUNC_COMPOSE_BASE_URL__ = Text.Combine(
                {
                    __VAR_ARRAY_CONECCTION_STRINGS__{0}, 
                    __VAR_ARRAY_CONECCTION_STRINGS__{2},
                    __VAR_ARRAY_CONECCTION_STRINGS__{1}
                }
            ),


            __VAR_DATASOURCE__ = OData.Feed(
                __FUNC_COMPOSE_BASE_URL__, null, 
                [
                    Implementation = __VAR_IMPLEMENTATION__
                ]
            ),


            __REMOVE_METADATA__ = Table.RemoveColumns(
                __VAR_DATASOURCE__, 
                __LIST_REMOVE_METADATA_USERS__

            ), 


            __ORDER_TABLE_USERS__ = Table.Sort(
                __REMOVE_METADATA__, 
                {
                    {
                        "Id", 
                        Order.Ascending
                    }
                }
            ),


            __ALTER_TYPE_ID_TEXT__ = Table.TransformColumnTypes(
                __ORDER_TABLE_USERS__,
                {
                    {
                        "UserPrincipalName", type text
                    }, 
                    {
                        "Id", type text
                    }
                }
            ), 


            __END_SISTEMA_X_USER_LIST__ = Table.SelectRows(
                __ALTER_TYPE_ID_TEXT__, each (
                    [UserPrincipalName] <> null
                )
            ),


        __VAR_DEFAULT_DOMAIN__ = "https://NOME_DA_EMPRESA.sharepoint.com/",
        __VAR_API_QUERYSTRING__ = "_api/web/siteusers",
        __VAR_STRING_SITE__ = "teams/SISTEMA_X/",
        __VAR_IMPLEMENTATION__ = "2.0"


    in __END_SISTEMA_X_USER_LIST__

    
in __VAR_SISTEMA_X_USER_LIST__
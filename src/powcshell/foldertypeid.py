"""既知フォルダ種類ID。"""

from comtypes import GUID


class FolderTypeID:
    """既知フォルダ種類ID定数。"""

    __slots__ = ()

    INVALID = GUID("{57807898-8c4f-4462-bb63-71042380b109}")
    GENERIC = GUID("{5c4f28b5-f869-4e84-8e60-f11db97c5cc7}")
    GENERIC_SEARCH_RESULTS = GUID("{7fde1a1e-8b31-49a5-93b8-6be14cfa4943}")
    GENERIC_LIBRARY = GUID("{5f4eab9a-6833-4f61-899d-31cf46979d49}")
    DOCUMENTS = GUID("{7d49d726-3c21-4f05-99aa-fdc2c9474656}")
    PICTURES = GUID("{b3690e58-e961-423b-b687-386ebfd83239}")
    MUSIC = GUID("{94d6ddcc-4a68-4175-a374-bd584a510b78}")
    VIDEOS = GUID("{5fa96407-7e77-483c-ac93-691d05850de8}")
    DOWNLOADS = GUID("{885a186e-a440-4ada-812b-db871b942259}")
    USERFILES = GUID("{CD0FC69B-71E2-46e5-9690-5BCD9F57AAB3}")
    USERS_LIBRARIES = GUID("{C4D98F09-6124-4fe0-9942-826416082DA9}")
    OTHERUSERS = GUID("{B337FD00-9DD5-4635-A6D4-DA33FD102B7A}")
    PUBLISHED_ITEMS = GUID("{7F2F5B96-FF74-41da-AFD8-1C78A5F3AEA2}")
    COMMUNICATIONS = GUID("{91475fe5-586b-4eba-8d75-d17434b8cdf6}")
    CONTACTS = GUID("{de2b70ec-9bf7-4a93-bd3d-243f7881d492}")
    STARTMENU = GUID("{ef87b4cb-f2ce-4785-8658-4ca6c63e38c6}")
    RECORDED_TV = GUID("{5557a28f-5da6-4f83-8809-c2c98a11a6fa}")
    SAVED_GAMES = GUID("{d0363307-28cb-4106-9f23-2956e3e5e0e7}")
    OPEN_SEARCH = GUID("{8faf9629-1980-46ff-8023-9dceab9c3ee3}")
    SEARCH_CONNECTOR = GUID("{982725ee-6f47-479e-b447-812bfa7d2e8f}")
    ACCOUNT_PICTURES = GUID("{db2a5d8f-06e6-4007-aba6-af877d526ea6}")
    GAMES_FOLDER = GUID("{b689b0d0-76d3-4cbb-87f7-585d0e0ce070}")
    CONTROLPANEL_CATEGORY = GUID("{de4f0660-fa10-4b8f-a494-068b20b22307}")
    CONTROLPANEL_CLASSIC = GUID("{0c3794f3-b545-43aa-a329-c37430c58d2a}")
    PRINTERS = GUID("{2c7bbec6-c844-4a0a-91fa-cef6f59cfda1}")
    RECYCLEBIN = GUID("{d6d9e004-cd87-442b-9d57-5e0aeb4f6f72}")
    SOFTWARE_EXPLORER = GUID("{d674391b-52d9-4e07-834e-67c98610f39d}")
    COMPRESSED_FOLDER = GUID("{80213e82-bcfd-4c4f-8817-bb27601267a9}")
    NETWORK_EXPLORER = GUID("{25CC242B-9A7C-4f51-80E0-7A2928FEBE42}")
    SEARCHES = GUID("{0b0ba2e3-405f-415e-a6ee-cad625207853}")
    SEARCH_HOME = GUID("{834d8a44-0974-4ed6-866e-f203d80b3810}")
    STORAGE_PROVIDER_GENERIC = GUID("{4F01EBC5-2385-41f2-A28E-2C5C91FB56E0}")
    STORAGE_PROVIDER_DOCUMENTS = GUID("{DD61BD66-70E8-48dd-9655-65C5E1AAC2D1}")
    STORAGE_PROVIDER_PICTURES = GUID("{71D642A9-F2B1-42cd-AD92-EB9300C7CC0A}")
    STORAGE_PROVIDER_MUSIC = GUID("{672ECD7E-AF04-4399-875C-0290845B6247}")
    STORAGE_PROVIDER_VIDEOS = GUID("{51294DA1-D7B1-485b-9E9A-17CFFE33E187}")
    VERSION_CONTROL = GUID("{69F1E26B-EC64-4280-BC83-F1EB887EC35A}")

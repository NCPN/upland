CREATE TABLE [tbl_SOP_version] (
  [version_key_number] LONG ,
  [SOP_number] LONG ,
  [SOP_version_number] VARCHAR ,
  [active_flag] VARCHAR ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([version_key_number], [SOP_number])
)

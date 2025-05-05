CREATE VIRTUAL TABLE documents USING fts5(
    name,
    text,
    extension,
    filepath,
    is_archived,
    archive_name
);

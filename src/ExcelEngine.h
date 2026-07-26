#pragma once

#include <QChar>
#include <QDateTime>
#include <QMap>
#include <QStringList>
#include <QVector>

enum class ChangeType { Added, Deleted, Modified, Unchanged };
enum class MergeJoin { FullOuter, LeftOnly, Inner };
enum class ConflictPolicy { PreferNew, PreferOld, MarkConflict };

struct Sheet {
    QString name;
    QStringList headers;
    QVector<QStringList> rows;
};

struct Workbook {
    QString path;
    QMap<QString, Sheet> sheets;
};

struct CellChange {
    QString column;
    QString oldValue;
    QString newValue;
};

struct Change {
    QString sheet;
    QString key;
    QString workbook1Key;
    QString workbook2Key;
    ChangeType type = ChangeType::Unchanged;
    QStringList headers;
    QStringList oldRow;
    QStringList newRow;
    QVector<CellChange> cells;
};

struct ColumnMapping {
    QString oldColumn;
    QString newColumn;
    QString outputColumn;
    bool key = false;
    bool compare = true;
    bool include = true;
};

struct Options {
    QVector<ColumnMapping> mappings;
    bool caseSensitive = false;
    bool trim = true;
    bool includeUnchanged = true;
};

struct Result {
    QString oldFile;
    QString newFile;
    QDateTime created;
    qsizetype added = 0;
    qsizetype deleted = 0;
    qsizetype modified = 0;
    qsizetype unchanged = 0;
    qsizetype workbook1Records = 0;
    qsizetype workbook2Records = 0;
    qsizetype totalRecords = 0;
    QStringList outputHeaders;
    QVector<Change> changes;
    QStringList warnings;
};

struct MergedSheet {
    QString name;
    QStringList headers;
    QVector<QStringList> rows;
};

QString typeName(ChangeType type);

class ExcelEngine
{
public:
    Workbook read(const QString& path) const;

    Result compareMapped(
        const Workbook& oldBook,
        const QString& oldSheetName,
        const Workbook& newBook,
        const QString& newSheetName,
        const Options& options) const;

    MergedSheet mergeMapped(
        const Workbook& oldBook,
        const QString& oldSheetName,
        const Workbook& newBook,
        const QString& newSheetName,
        const Options& options,
        MergeJoin join,
        ConflictPolicy conflictPolicy) const;

    static void exportXlsx(const Result& result, const QString& path);
    static void exportCsv(const Result& result, const QString& path);
    static void exportPdf(const Result& result, const QString& path);
    static void exportMergedXlsx(const MergedSheet& sheet, const QString& path);

private:
    Workbook readXlsx(const QString& path, const QString& originalPath = {}) const;
    Workbook readXls(const QString& path) const;
    Workbook readDelimited(const QString& path, QChar delimiter) const;
    Workbook convertAndRead(const QString& path) const;
    QString findLibreOffice() const;
};

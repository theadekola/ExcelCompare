#include "ExcelEngine.h"

#include "xlsxcellrange.h"
#include "xlsxdocument.h"
#include "xlsxformat.h"

#include <xls.h>

using namespace xls;

#include <QColor>
#include <QDate>
#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QHash>
#include <QPageLayout>
#include <QPageSize>
#include <QPainter>
#include <QPdfWriter>
#include <QProcess>
#include <QRegularExpression>
#include <QSet>
#include <QStandardPaths>
#include <QTemporaryDir>
#include <QTextStream>

#include <stdexcept>

namespace {

QString normalise(QString value, const Options& options)
{
    if (options.trim)
        value = value.trimmed();
    return options.caseSensitive ? value : value.toCaseFolded();
}

QString cellValue(const Sheet& sheet, const QStringList& row, const QString& header)
{
    if (header.isEmpty())
        return {};
    const int index = sheet.headers.indexOf(header);
    return index >= 0 && index < row.size() ? row.at(index) : QString();
}

QString csvQuote(QString text)
{
    text.replace('"', "\"\"");
    return "\"" + text + "\"";
}

QString sourceDisplayText(QString text)
{
    const QString trimmed = text.trimmed();
    static const QRegularExpression decimalPattern(QStringLiteral("^[+-]?\\d+\\.\\d+$"));
    static const QRegularExpression leadingZeroInteger(QStringLiteral("^[+-]?0\\d+$"));
    if (!decimalPattern.match(trimmed).hasMatch() || leadingZeroInteger.match(trimmed).hasMatch())
        return text;
    QString cleaned = trimmed;
    while (cleaned.endsWith('0'))
        cleaned.chop(1);
    if (cleaned.endsWith('.'))
        cleaned.chop(1);
    return cleaned;
}

QString resultDisplayValue(const Change& change, int field)
{
    const QString oldValue = field < change.oldRow.size() ? change.oldRow.at(field) : QString();
    const QString newValue = field < change.newRow.size() ? change.newRow.at(field) : QString();
    if (change.type == ChangeType::Deleted)
        return oldValue;
    if (change.type == ChangeType::Modified && oldValue != newValue)
        return QString("%1 -> %2").arg(oldValue, newValue);
    return newValue.isEmpty() ? oldValue : newValue;
}



bool isIdentifierHeader(const QString& header)
{
    const QString h = header.trimmed().toCaseFolded();
    return h.contains("key") || h.contains(" id") || h.endsWith("id")
        || h.contains("no") || h.contains("number") || h.contains("code")
        || h.contains("form") || h.contains("account") || h.contains("phone")
        || h.contains("reference") || h.contains("ref");
}

bool isDateHeader(const QString& header)
{
    const QString h = header.trimmed().toCaseFolded();
    return h == "date" || h.contains(" date") || h.endsWith("date");
}

void writeExcelValue(QXlsx::Document& document,
                     int row,
                     int column,
                     const QString& header,
                     const QString& sourceValue,
                     const QXlsx::Format& baseFormat,
                     bool forceText = false)
{
    const QString value = sourceValue;
    const QString trimmed = value.trimmed();
    if (trimmed.isEmpty()) {
        document.write(row, column, QString(), baseFormat);
        return;
    }

    // Values that combine old/new text are report descriptions, not numbers.
    if (forceText || trimmed.contains(" -> ") || trimmed.contains("[Workbook")) {
        document.write(row, column, value, baseFormat);
        return;
    }

    static const QRegularExpression integerPattern(QStringLiteral("^[+-]?\\d+$"));
    static const QRegularExpression decimalPattern(QStringLiteral("^[+-]?(?:\\d+\\.\\d+|\\d+)$"));
    const bool integerLike = integerPattern.match(trimmed).hasMatch();

    // Write Excel serial dates as real date cells. This removes the
    // "Number Stored as Text" warning and displays a readable date.
    if (isDateHeader(header) && integerLike) {
        bool ok = false;
        const qlonglong serial = trimmed.toLongLong(&ok);
        if (ok && serial > 0 && serial < 300000) {
            QXlsx::Format dateFormat = baseFormat;
            dateFormat.setNumberFormat("dd/mm/yyyy");
            document.write(row, column, QDate(1899, 12, 30).addDays(serial), dateFormat);
            return;
        }
    }

    // Identifier columns remain visually exact. Pure numeric identifiers are
    // stored as numbers. Leading zeros are restored with a custom number mask,
    // so Excel does not show the green "Number Stored as Text" warning.
    if (isIdentifierHeader(header) && integerLike) {
        bool ok = false;
        const qlonglong numeric = trimmed.toLongLong(&ok);
        if (ok) {
            QXlsx::Format idFormat = baseFormat;
            QString digits = trimmed;
            if (digits.startsWith('+') || digits.startsWith('-'))
                digits.remove(0, 1);
            if (digits.size() > 1 && digits.startsWith('0'))
                idFormat.setNumberFormat(QString(digits.size(), '0'));
            else
                idFormat.setNumberFormat("0");
            document.write(row, column, numeric, idFormat);
            return;
        }
    }

    // Ordinary numeric data is written as a real numeric cell.
    if (decimalPattern.match(trimmed).hasMatch()) {
        bool ok = false;
        const double numeric = trimmed.toDouble(&ok);
        if (ok) {
            QXlsx::Format numberFormat = baseFormat;
            const int decimalPoint = trimmed.indexOf('.');
            if (decimalPoint < 0 || trimmed.mid(decimalPoint + 1).remove('0').isEmpty())
                numberFormat.setNumberFormat("#,##0");
            else
                numberFormat.setNumberFormat("#,##0.########");
            document.write(row, column, numeric, numberFormat);
            return;
        }
    }

    document.write(row, column, value, baseFormat);
}

QXlsx::Format headerFormat()
{
    QXlsx::Format format;
    format.setFontBold(true);
    format.setFontColor(Qt::white);
    format.setPatternBackgroundColor(QColor("#1e3a8a"));
    format.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    format.setVerticalAlignment(QXlsx::Format::AlignVCenter);
    return format;
}

QXlsx::Format statusFormat(ChangeType type)
{
    QXlsx::Format format;
    switch (type) {
    case ChangeType::Added:
        format.setPatternBackgroundColor(QColor("#dcfce7"));
        format.setFontColor(QColor("#166534"));
        break;
    case ChangeType::Deleted:
        format.setPatternBackgroundColor(QColor("#fee2e2"));
        format.setFontColor(QColor("#991b1b"));
        break;
    case ChangeType::Modified:
        format.setPatternBackgroundColor(QColor("#fef3c7"));
        format.setFontColor(QColor("#92400e"));
        break;
    case ChangeType::Unchanged:
        format.setPatternBackgroundColor(QColor("#f1f5f9"));
        format.setFontColor(QColor("#475569"));
        break;
    }
    return format;
}

QStringList parseDelimitedLine(const QString& line, QChar delimiter)
{
    QStringList fields;
    QString field;
    bool insideQuotes = false;

    for (qsizetype index = 0; index < line.size(); ++index) {
        const QChar character = line.at(index);
        if (character == '"') {
            if (insideQuotes && index + 1 < line.size() && line.at(index + 1) == '"') {
                field.append('"');
                ++index;
            } else {
                insideQuotes = !insideQuotes;
            }
        } else if (character == delimiter && !insideQuotes) {
            fields.append(field);
            field.clear();
        } else {
            field.append(character);
        }
    }

    fields.append(field);
    return fields;
}

QString makeUniqueHeader(const QString& original, int columnNumber, QSet<QString>& used)
{
    QString header = original.trimmed();
    if (header.isEmpty())
        header = QString("Column %1").arg(columnNumber);

    QString unique = header;
    int suffix = 2;
    while (used.contains(unique))
        unique = QString("%1 (%2)").arg(header).arg(suffix++);

    used.insert(unique);
    return unique;
}

QVector<ColumnMapping> comparisonMappings(const Options& options)
{
    QVector<ColumnMapping> output;
    for (const ColumnMapping& mapping : options.mappings) {
        if (mapping.compare && (!mapping.oldColumn.isEmpty() || !mapping.newColumn.isEmpty()))
            output.append(mapping);
    }
    return output;
}

QVector<ColumnMapping> mergeMappings(const Options& options)
{
    QVector<ColumnMapping> output;
    for (const ColumnMapping& mapping : options.mappings) {
        if (mapping.include && (!mapping.oldColumn.isEmpty() || !mapping.newColumn.isEmpty()))
            output.append(mapping);
    }
    return output;
}

QVector<ColumnMapping> keyMappings(const Options& options)
{
    QVector<ColumnMapping> output;
    for (const ColumnMapping& mapping : options.mappings) {
        // A key identifies records and is independent of whether the column is
        // selected for comparison output or merged output.
        if (mapping.key && !mapping.oldColumn.isEmpty() && !mapping.newColumn.isEmpty())
            output.append(mapping);
    }
    return output;
}

QString buildKey(const Sheet& sheet,
                 const QStringList& row,
                 const QVector<ColumnMapping>& keys,
                 bool oldSide,
                 const Options& options,
                 bool& valid)
{
    valid = true;
    QStringList parts;

    for (const ColumnMapping& mapping : keys) {
        const QString header = oldSide ? mapping.oldColumn : mapping.newColumn;
        const QString value = normalise(cellValue(sheet, row, header), options);
        if (value.isEmpty()) {
            valid = false;
            return {};
        }
        parts.append(value);
    }

    if (parts.isEmpty()) {
        valid = false;
        return {};
    }

    return parts.join(QChar(31));
}

QString buildDisplayKey(const Sheet& sheet,
                        const QStringList& row,
                        const QVector<ColumnMapping>& keys,
                        bool oldSide,
                        bool& valid)
{
    valid = true;
    QStringList parts;

    for (const ColumnMapping& mapping : keys) {
        const QString header = oldSide ? mapping.oldColumn : mapping.newColumn;
        const QString value = cellValue(sheet, row, header);
        if (value.trimmed().isEmpty()) {
            valid = false;
            return {};
        }
        // Keep the exact source representation for results and exports.
        // Matching may use a normalised copy, but prefixes, leading zeroes,
        // letter case and source spacing are never rewritten here.
        parts.append(value);
    }

    if (parts.isEmpty()) {
        valid = false;
        return {};
    }

    return parts.join(" | ");
}

QString outputName(const ColumnMapping& mapping, int index)
{
    if (!mapping.outputColumn.trimmed().isEmpty())
        return mapping.outputColumn.trimmed();
    if (!mapping.newColumn.isEmpty())
        return mapping.newColumn;
    if (!mapping.oldColumn.isEmpty())
        return mapping.oldColumn;
    return QString("Column %1").arg(index + 1);
}

} // namespace

QString typeName(ChangeType type)
{
    switch (type) {
    case ChangeType::Added: return "➕ New Record";
    case ChangeType::Deleted: return "⚠ Unmatched";
    case ChangeType::Modified: return "✏ Modified";
    case ChangeType::Unchanged: return "✓ Matched";
    }
    return "Unknown";
}

Workbook ExcelEngine::read(const QString& path) const
{
    if (!QFileInfo::exists(path))
        throw std::runtime_error("The selected spreadsheet does not exist.");

    const QString extension = QFileInfo(path).suffix().toLower();
    if (extension == "xlsx" || extension == "xlsm" || extension == "xltx" || extension == "xltm")
        return readXlsx(path);
    if (extension == "xls")
        return readXls(path);
    if (extension == "csv")
        return readDelimited(path, ',');
    if (extension == "tsv" || extension == "tab")
        return readDelimited(path, '\t');

    const QStringList convertibleFormats = {
        "xlsb", "xlt", "ods", "fods"
    };
    if (convertibleFormats.contains(extension))
        return convertAndRead(path);

    throw std::runtime_error(
        QString("Unsupported spreadsheet format '.%1'. Supported formats: "
                ".xlsx, .xls, .xlsb, .xlsm, .xlt, .xltx, .xltm, .ods, .fods, .csv and .tsv.")
            .arg(extension).toStdString());
}

Workbook ExcelEngine::readXlsx(const QString& path, const QString& originalPath) const
{
    QXlsx::Document document(path);
    if (!document.load())
        throw std::runtime_error("The spreadsheet could not be opened as an Excel workbook.");

    Workbook output;
    output.path = originalPath.isEmpty() ? path : originalPath;

    for (const QString& sheetName : document.sheetNames()) {
        if (!document.selectSheet(sheetName))
            continue;

        Sheet sheet;
        sheet.name = sheetName;
        const QXlsx::CellRange dimension = document.dimension();

        if (!dimension.isValid()) {
            output.sheets.insert(sheetName, sheet);
            continue;
        }

        QSet<QString> usedHeaders;
        for (int column = dimension.firstColumn(); column <= dimension.lastColumn(); ++column) {
            const QString rawHeader = document.read(dimension.firstRow(), column).toString();
            sheet.headers.append(makeUniqueHeader(rawHeader, column, usedHeaders));
        }

        for (int rowNumber = dimension.firstRow() + 1; rowNumber <= dimension.lastRow(); ++rowNumber) {
            QStringList row;
            bool containsValue = false;
            for (int column = dimension.firstColumn(); column <= dimension.lastColumn(); ++column) {
                const QString value = document.read(rowNumber, column).toString();
                row.append(value);
                if (!value.trimmed().isEmpty())
                    containsValue = true;
            }
            if (containsValue)
                sheet.rows.append(row);
        }

        output.sheets.insert(sheetName, sheet);
    }

    return output;
}

Workbook ExcelEngine::readXls(const QString& path) const
{
    const QByteArray nativePath = QFile::encodeName(QDir::toNativeSeparators(path));
    xls_error_t error = LIBXLS_OK;
    xlsWorkBook* workbook = xls_open_file(nativePath.constData(), "UTF-8", &error);

    if (!workbook) {
        throw std::runtime_error(
            QString("Cannot open legacy Excel file '%1'. libxls error: %2")
                .arg(QFileInfo(path).fileName(), QString::fromUtf8(xls_getError(error)))
                .toStdString());
    }

    struct WorkbookGuard {
        xlsWorkBook* value;
        ~WorkbookGuard() { if (value) xls_close_WB(value); }
    } guard{workbook};

    Workbook output;
    output.path = path;

    for (int sheetIndex = 0; sheetIndex < workbook->sheets.count; ++sheetIndex) {
        xlsWorkSheet* worksheet = xls_getWorkSheet(workbook, sheetIndex);
        if (!worksheet)
            continue;

        struct WorksheetGuard {
            xlsWorkSheet* value;
            ~WorksheetGuard() { if (value) xls_close_WS(value); }
        } worksheetGuard{worksheet};

        error = xls_parseWorkSheet(worksheet);
        if (error != LIBXLS_OK) {
            throw std::runtime_error(
                QString("Unable to parse worksheet %1 in '%2': %3")
                    .arg(sheetIndex + 1)
                    .arg(QFileInfo(path).fileName(), QString::fromUtf8(xls_getError(error)))
                    .toStdString());
        }

        QString sheetName;
        if (sheetIndex < workbook->sheets.count && workbook->sheets.sheet[sheetIndex].name)
            sheetName = QString::fromUtf8(workbook->sheets.sheet[sheetIndex].name);
        if (sheetName.trimmed().isEmpty())
            sheetName = QString("Sheet%1").arg(sheetIndex + 1);

        Sheet sheet;
        sheet.name = sheetName;

        const int lastRow = static_cast<int>(worksheet->rows.lastrow);
        const int lastColumn = static_cast<int>(worksheet->rows.lastcol);
        if (lastRow < 0 || lastColumn < 0) {
            output.sheets.insert(sheetName, sheet);
            continue;
        }

        auto textForCell = [](const xlsCell* cell) -> QString {
            if (!cell || cell->id == XLS_RECORD_BLANK)
                return {};

            if (cell->id == XLS_RECORD_NUMBER || cell->id == XLS_RECORD_RK) {
                // libxls may provide the formatted display text in str. Prefer it
                // so identifiers such as 0000057 and other formatted values are
                // preserved exactly as shown in the source workbook.
                if (cell->str && *cell->str)
                    return sourceDisplayText(QString::fromUtf8(cell->str));
                return sourceDisplayText(QString::number(cell->d, 'g', 15));
            }

            if (cell->id == XLS_RECORD_FORMULA) {
                if (cell->str) {
                    const QString formulaResult = QString::fromUtf8(cell->str);
                    if (formulaResult.compare("bool", Qt::CaseInsensitive) == 0)
                        return cell->d != 0.0 ? "TRUE" : "FALSE";
                    if (formulaResult.compare("error", Qt::CaseInsensitive) == 0)
                        return "#ERROR";
                    if (!formulaResult.isEmpty())
                        return sourceDisplayText(formulaResult);
                }
                return sourceDisplayText(QString::number(cell->d, 'g', 15));
            }

            if (cell->str)
                return sourceDisplayText(QString::fromUtf8(cell->str));

            return sourceDisplayText(QString::number(cell->d, 'g', 15));
        };

        xlsRow* headerRow = xls_row(worksheet, 0);
        QSet<QString> usedHeaders;
        for (int column = 0; column <= lastColumn; ++column) {
            const xlsCell* cell = headerRow ? &headerRow->cells.cell[column] : nullptr;
            sheet.headers.append(makeUniqueHeader(textForCell(cell), column + 1, usedHeaders));
        }

        for (int rowNumber = 1; rowNumber <= lastRow; ++rowNumber) {
            xlsRow* sourceRow = xls_row(worksheet, rowNumber);
            QStringList row;
            bool containsValue = false;

            for (int column = 0; column <= lastColumn; ++column) {
                const xlsCell* cell = sourceRow ? &sourceRow->cells.cell[column] : nullptr;
                const QString value = textForCell(cell);
                row.append(value);
                if (!value.trimmed().isEmpty())
                    containsValue = true;
            }

            if (containsValue)
                sheet.rows.append(row);
        }

        QString uniqueSheetName = sheetName;
        int duplicate = 2;
        while (output.sheets.contains(uniqueSheetName))
            uniqueSheetName = QString("%1 (%2)").arg(sheetName).arg(duplicate++);
        sheet.name = uniqueSheetName;
        output.sheets.insert(uniqueSheetName, sheet);
    }

    if (output.sheets.isEmpty())
        throw std::runtime_error("The .xls workbook contains no readable worksheets.");

    return output;
}

Workbook ExcelEngine::readDelimited(const QString& path, QChar delimiter) const
{
    QFile file(path);
    if (!file.open(QIODevice::ReadOnly | QIODevice::Text))
        throw std::runtime_error("The delimited spreadsheet could not be opened.");

    QTextStream stream(&file);
    stream.setEncoding(QStringConverter::Utf8);

    Workbook output;
    output.path = path;
    Sheet sheet;
    sheet.name = QFileInfo(path).completeBaseName();
    if (sheet.name.trimmed().isEmpty())
        sheet.name = "Data";

    if (stream.atEnd()) {
        output.sheets.insert(sheet.name, sheet);
        return output;
    }

    QString firstLine = stream.readLine();
    if (!firstLine.isEmpty() && firstLine.front() == QChar(0xFEFF))
        firstLine.removeFirst();

    sheet.headers = parseDelimitedLine(firstLine, delimiter);
    QSet<QString> usedHeaders;
    for (int index = 0; index < sheet.headers.size(); ++index)
        sheet.headers[index] = makeUniqueHeader(sheet.headers.at(index), index + 1, usedHeaders);

    while (!stream.atEnd()) {
        const QString line = stream.readLine();
        if (line.trimmed().isEmpty())
            continue;

        QStringList row = parseDelimitedLine(line, delimiter);
        while (row.size() < sheet.headers.size())
            row.append(QString());
        while (row.size() > sheet.headers.size())
            sheet.headers.append(QString("Column %1").arg(sheet.headers.size() + 1));
        sheet.rows.append(row);
    }

    output.sheets.insert(sheet.name, sheet);
    return output;
}

QString ExcelEngine::findLibreOffice() const
{
    QString executable = QStandardPaths::findExecutable("soffice");
    if (!executable.isEmpty())
        return executable;
    executable = QStandardPaths::findExecutable("soffice.exe");
    if (!executable.isEmpty())
        return executable;

    const QStringList possiblePaths = {
        "C:/Program Files/LibreOffice/program/soffice.exe",
        "C:/Program Files (x86)/LibreOffice/program/soffice.exe"
    };
    for (const QString& path : possiblePaths) {
        if (QFileInfo::exists(path))
            return path;
    }
    return {};
}

Workbook ExcelEngine::convertAndRead(const QString& path) const
{
    const QString libreOffice = findLibreOffice();
    if (libreOffice.isEmpty()) {
        throw std::runtime_error(
            "This file format requires LibreOffice. Install LibreOffice and try again. "
            "The source file is only converted temporarily and is not modified.");
    }

    QTemporaryDir temporaryDirectory;
    if (!temporaryDirectory.isValid())
        throw std::runtime_error("A temporary conversion folder could not be created.");

    QProcess process;
    const QStringList arguments = {
        "--headless", "--convert-to", "xlsx", "--outdir",
        temporaryDirectory.path(), QDir::toNativeSeparators(path)
    };

    process.start(libreOffice, arguments);
    if (!process.waitForStarted(15000))
        throw std::runtime_error("LibreOffice could not be started.");
    if (!process.waitForFinished(120000)) {
        process.kill();
        process.waitForFinished();
        throw std::runtime_error("Spreadsheet conversion timed out.");
    }

    QString convertedFile = QDir(temporaryDirectory.path()).filePath(
        QFileInfo(path).completeBaseName() + ".xlsx");

    if (!QFileInfo::exists(convertedFile)) {
        const QStringList convertedFiles = QDir(temporaryDirectory.path()).entryList({"*.xlsx"}, QDir::Files);
        if (!convertedFiles.isEmpty())
            convertedFile = QDir(temporaryDirectory.path()).filePath(convertedFiles.first());
    }

    if (!QFileInfo::exists(convertedFile)) {
        const QString output = QString::fromLocal8Bit(process.readAllStandardOutput());
        const QString error = QString::fromLocal8Bit(process.readAllStandardError());
        throw std::runtime_error(
            QString("LibreOffice could not convert the spreadsheet.\nOutput: %1\nError: %2")
                .arg(output, error).toStdString());
    }

    return readXlsx(convertedFile, path);
}

Result ExcelEngine::compareMapped(const Workbook& oldBook,
                                  const QString& oldSheetName,
                                  const Workbook& newBook,
                                  const QString& newSheetName,
                                  const Options& options) const
{
    const QVector<ColumnMapping> mappings = comparisonMappings(options);
    const QVector<ColumnMapping> keys = keyMappings(options);
    if (keys.isEmpty())
        throw std::runtime_error("Select at least one mapped key column in both workbooks.");
    if (mappings.isEmpty())
        throw std::runtime_error("Tick Compare for at least one mapped column.");
    if (!oldBook.sheets.contains(oldSheetName) || !newBook.sheets.contains(newSheetName))
        throw std::runtime_error("Select a valid worksheet from both workbooks.");

    const Sheet oldSheet = oldBook.sheets.value(oldSheetName);
    const Sheet newSheet = newBook.sheets.value(newSheetName);

    Result result;
    result.oldFile = oldBook.path;
    result.newFile = newBook.path;
    result.created = QDateTime::currentDateTime();
    result.workbook1Records = oldSheet.rows.size();
    result.workbook2Records = newSheet.rows.size();
    result.totalRecords = result.workbook1Records;
    for (int i = 0; i < mappings.size(); ++i)
        result.outputHeaders.append(outputName(mappings.at(i), i));

    QHash<QString, QStringList> oldRecords;
    QHash<QString, QStringList> newRecords;
    QHash<QString, QString> oldDisplayKeys;
    QHash<QString, QString> newDisplayKeys;

    auto indexRows = [&](const Sheet& sheet, bool oldSide,
                         QHash<QString, QStringList>& records,
                         QHash<QString, QString>& displayKeys,
                         const QString& sideName) {
        int rowNumber = 2;
        QHash<QString, int> duplicateCounts;
        for (const QStringList& row : sheet.rows) {
            bool valid = false;
            QString key = buildKey(sheet, row, keys, oldSide, options, valid);
            bool displayValid = false;
            QString displayKey = buildDisplayKey(sheet, row, keys, oldSide, displayValid);
            if (!valid) {
                // Never discard a source row. A row without a usable key is assigned a
                // side-specific synthetic key so it remains visible in results and exports.
                key = QString("Missing key - %1 row %2").arg(sideName).arg(rowNumber);
                displayKey = key;
                result.warnings.append(QString("%1 row %2 has a missing key and was preserved as a separate record.")
                                           .arg(sideName).arg(rowNumber));
            } else if (records.contains(key)) {
                // Preserve duplicate-key rows instead of replacing an earlier row.
                const int duplicateNumber = ++duplicateCounts[key];
                const QString originalKey = key;
                key = QString("%1 [duplicate %2 - %3 row %4]")
                          .arg(originalKey).arg(duplicateNumber).arg(sideName).arg(rowNumber);
                result.warnings.append(QString("%1 contains duplicate key '%2' at row %3. The row was preserved separately.")
                                           .arg(sideName, originalKey).arg(rowNumber));
            }
            records.insert(key, row);
            displayKeys.insert(key, displayValid ? displayKey : key);
            ++rowNumber;
        }
    };

    indexRows(oldSheet, true, oldRecords, oldDisplayKeys, "Workbook 1");
    indexRows(newSheet, false, newRecords, newDisplayKeys, "Workbook 2");

    QSet<QString> allKeys;
    for (auto it = oldRecords.cbegin(); it != oldRecords.cend(); ++it)
        allKeys.insert(it.key());
    for (auto it = newRecords.cbegin(); it != newRecords.cend(); ++it)
        allKeys.insert(it.key());

    QStringList sortedKeys = allKeys.values();
    sortedKeys.sort(options.caseSensitive ? Qt::CaseSensitive : Qt::CaseInsensitive);

    for (const QString& key : sortedKeys) {
        const bool hasOld = oldRecords.contains(key);
        const bool hasNew = newRecords.contains(key);
        Change change;
        change.sheet = QString("%1 ↔ %2").arg(oldSheetName, newSheetName);
        // Use the original source key in the UI and every export. The normalised
        // key is used only internally for record matching.
        change.workbook1Key = hasOld ? oldDisplayKeys.value(key, key) : QString();
        change.workbook2Key = hasNew ? newDisplayKeys.value(key, key) : QString();
        // Record key always uses the exact source text. Workbook 1 is the
        // reference side in compare-only mode; new-only records use Workbook 2.
        change.key = hasOld ? change.workbook1Key : change.workbook2Key;

        for (int i = 0; i < mappings.size(); ++i)
            change.headers.append(outputName(mappings.at(i), i));

        if (!hasOld) {
            change.type = ChangeType::Added;
            for (const ColumnMapping& mapping : mappings) {
                change.oldRow.append(QString());
                change.newRow.append(cellValue(newSheet, newRecords.value(key), mapping.newColumn));
            }
            ++result.added;
        } else if (!hasNew) {
            change.type = ChangeType::Deleted;
            for (const ColumnMapping& mapping : mappings) {
                change.oldRow.append(cellValue(oldSheet, oldRecords.value(key), mapping.oldColumn));
                change.newRow.append(QString());
            }
            ++result.deleted;
        } else {
            const QStringList oldRow = oldRecords.value(key);
            const QStringList newRow = newRecords.value(key);
            for (int i = 0; i < mappings.size(); ++i) {
                const ColumnMapping& mapping = mappings.at(i);
                const QString oldValue = cellValue(oldSheet, oldRow, mapping.oldColumn);
                const QString newValue = cellValue(newSheet, newRow, mapping.newColumn);
                change.oldRow.append(oldValue);
                change.newRow.append(newValue);
                if (mapping.compare && !mapping.oldColumn.isEmpty() && !mapping.newColumn.isEmpty() &&
                    normalise(oldValue, options) != normalise(newValue, options)) {
                    change.cells.append({outputName(mapping, i), oldValue, newValue});
                }
            }
            change.type = change.cells.isEmpty() ? ChangeType::Unchanged : ChangeType::Modified;
            if (change.type == ChangeType::Modified)
                ++result.modified;
            else
                ++result.unchanged;
        }

        // Keep every comparison outcome in memory so the Results filter can
        // display Added, Deleted, Modified and Unchanged records independently.
        result.changes.append(change);
    }

    return result;
}

MergedSheet ExcelEngine::mergeMapped(const Workbook& oldBook,
                                     const QString& oldSheetName,
                                     const Workbook& newBook,
                                     const QString& newSheetName,
                                     const Options& options,
                                     MergeJoin join,
                                     ConflictPolicy conflictPolicy) const
{
    const QVector<ColumnMapping> mappings = mergeMappings(options);
    const QVector<ColumnMapping> keys = keyMappings(options);
    if (keys.isEmpty())
        throw std::runtime_error("Select at least one mapped key column before merging.");
    if (mappings.isEmpty())
        throw std::runtime_error("Select at least one output column before merging.");
    if (!oldBook.sheets.contains(oldSheetName) || !newBook.sheets.contains(newSheetName))
        throw std::runtime_error("Select a valid worksheet from both workbooks.");

    const Sheet oldSheet = oldBook.sheets.value(oldSheetName);
    const Sheet newSheet = newBook.sheets.value(newSheetName);
    QHash<QString, QStringList> oldRecords;
    QHash<QString, QStringList> newRecords;

    auto indexRows = [&](const Sheet& sheet, bool oldSide, QHash<QString, QStringList>& records) {
        int rowNumber = 2;
        QHash<QString, int> duplicateCounts;
        const QString sideName = oldSide ? "Workbook 1" : "Workbook 2";
        for (const QStringList& row : sheet.rows) {
            bool valid = false;
            QString key = buildKey(sheet, row, keys, oldSide, options, valid);
            if (!valid)
                key = QString("Missing key - %1 row %2").arg(sideName).arg(rowNumber);
            else if (records.contains(key)) {
                const int duplicateNumber = ++duplicateCounts[key];
                key = QString("%1 [duplicate %2 - %3 row %4]")
                          .arg(key).arg(duplicateNumber).arg(sideName).arg(rowNumber);
            }
            records.insert(key, row);
            ++rowNumber;
        }
    };
    indexRows(oldSheet, true, oldRecords);
    indexRows(newSheet, false, newRecords);

    QSet<QString> keysToMerge;
    if (join == MergeJoin::LeftOnly || join == MergeJoin::FullOuter) {
        for (auto it = oldRecords.cbegin(); it != oldRecords.cend(); ++it)
            keysToMerge.insert(it.key());
    }
    if (join == MergeJoin::FullOuter) {
        for (auto it = newRecords.cbegin(); it != newRecords.cend(); ++it)
            keysToMerge.insert(it.key());
    }
    if (join == MergeJoin::Inner) {
        for (auto it = oldRecords.cbegin(); it != oldRecords.cend(); ++it) {
            if (newRecords.contains(it.key()))
                keysToMerge.insert(it.key());
        }
    }

    QStringList sortedKeys = keysToMerge.values();
    sortedKeys.sort(options.caseSensitive ? Qt::CaseSensitive : Qt::CaseInsensitive);

    MergedSheet merged;
    merged.name = "Merged Data";
    merged.headers.append("Merge Status");
    for (int i = 0; i < mappings.size(); ++i)
        merged.headers.append(outputName(mappings.at(i), i));

    for (const QString& key : sortedKeys) {
        const bool hasOld = oldRecords.contains(key);
        const bool hasNew = newRecords.contains(key);
        QStringList outputRow;
        outputRow.append(hasOld && hasNew ? "Matched" : hasNew ? "Workbook 2 only" : "Workbook 1 only");

        for (const ColumnMapping& mapping : mappings) {
            const QString oldValue = hasOld ? cellValue(oldSheet, oldRecords.value(key), mapping.oldColumn) : QString();
            const QString newValue = hasNew ? cellValue(newSheet, newRecords.value(key), mapping.newColumn) : QString();
            QString selected;

            if (!hasOld)
                selected = newValue;
            else if (!hasNew)
                selected = oldValue;
            else if (oldValue.isEmpty())
                selected = newValue;
            else if (newValue.isEmpty())
                selected = oldValue;
            else if (normalise(oldValue, options) == normalise(newValue, options))
                selected = newValue;
            else if (conflictPolicy == ConflictPolicy::PreferOld)
                selected = oldValue;
            else if (conflictPolicy == ConflictPolicy::PreferNew)
                selected = newValue;
            else
                selected = QString("[Workbook 1: %1] [Workbook 2: %2]").arg(oldValue, newValue);

            outputRow.append(selected);
        }
        merged.rows.append(outputRow);
    }

    return merged;
}

void ExcelEngine::exportCsv(const Result& result, const QString& path)
{
    QFile file(path);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Text))
        throw std::runtime_error("Cannot create CSV report.");

    QTextStream stream(&file);
    stream.setEncoding(QStringConverter::Utf8);

    stream << "Worksheet pair,Status,Record key,Workbook 1 key,Workbook 2 key";
    for (const QString& header : result.outputHeaders)
        stream << ',' << csvQuote(header);
    stream << ",Changed columns,Timestamp\n";

    for (const Change& change : result.changes) {
        stream << csvQuote(change.sheet) << ','
               << csvQuote(typeName(change.type)) << ','
               << csvQuote(change.key) << ','
               << csvQuote(change.workbook1Key) << ','
               << csvQuote(change.workbook2Key);

        for (int field = 0; field < result.outputHeaders.size(); ++field) {
            stream << ',' << csvQuote(resultDisplayValue(change, field));
        }

        QStringList changedColumns;
        for (const CellChange& cell : change.cells)
            changedColumns.append(cell.column);

        stream << ',' << csvQuote(changedColumns.join(", "))
               << ',' << csvQuote(result.created.toString(Qt::ISODate)) << '\n';
    }
}

void ExcelEngine::exportXlsx(const Result& result, const QString& path)
{
    QXlsx::Document document;
    document.renameSheet("Sheet1", "Comparison Results");

    QStringList resultHeaders = {"Worksheet pair", "Status", "Record key", "Workbook 1 key", "Workbook 2 key"};
    resultHeaders.append(result.outputHeaders);
    resultHeaders.append("Changed columns");
    for (int column = 0; column < resultHeaders.size(); ++column)
        document.write(1, column + 1, resultHeaders.at(column), headerFormat());

    int row = 2;
    for (const Change& change : result.changes) {
        const QXlsx::Format rowFormat = statusFormat(change.type);
        writeExcelValue(document, row, 1, "Worksheet pair", change.sheet, rowFormat, true);
        writeExcelValue(document, row, 2, "Status", typeName(change.type), rowFormat, true);
        writeExcelValue(document, row, 3, "Record key", change.key, rowFormat);
        writeExcelValue(document, row, 4, "Workbook 1 key", change.workbook1Key, rowFormat);
        writeExcelValue(document, row, 5, "Workbook 2 key", change.workbook2Key, rowFormat);
        for (int field = 0; field < result.outputHeaders.size(); ++field)
            writeExcelValue(document, row, 6 + field, result.outputHeaders.at(field), resultDisplayValue(change, field), rowFormat);
        QStringList changedColumns;
        for (const CellChange& cell : change.cells)
            changedColumns.append(cell.column);
        document.write(row, 6 + result.outputHeaders.size(), changedColumns.join(", "), rowFormat);
        ++row;
    }

    document.setColumnWidth(1, 1, 20);
    document.setColumnWidth(2, 2, 18);
    document.setColumnWidth(3, 5, 24);
    for (int column = 6; column < 6 + result.outputHeaders.size(); ++column)
        document.setColumnWidth(column, column, 24);
    document.setColumnWidth(6 + result.outputHeaders.size(), 6 + result.outputHeaders.size(), 30);

    document.addSheet("Summary");
    document.selectSheet("Summary");
    document.write("A1", "Excel Compare Professional Report", headerFormat());
    document.write("A2", "Generated by AAT-Tech Ltd");
    document.write("A4", "Workbook 1"); document.write("B4", result.oldFile);
    document.write("A5", "Workbook 2"); document.write("B5", result.newFile);
    document.write("A6", "Generated"); document.write("B6", result.created.toString(Qt::ISODate));
    document.write("A8", "New Records"); document.write("B8", static_cast<qlonglong>(result.added));
    document.write("A9", "Unmatched"); document.write("B9", static_cast<qlonglong>(result.deleted));
    document.write("A10", "Modified"); document.write("B10", static_cast<qlonglong>(result.modified));
    document.write("A11", "Matched"); document.write("B11", static_cast<qlonglong>(result.unchanged));
    document.write("A12", "Total Records"); document.write("B12", static_cast<qlonglong>(result.totalRecords));
    document.setColumnWidth(1, 1, 24);
    document.setColumnWidth(2, 2, 70);

    document.addSheet("Modified Field Details");
    document.selectSheet("Modified Field Details");
    const QStringList detailHeaders = {"Worksheet", "Status", "Record Key", "Workbook 1 Key", "Workbook 2 Key", "Column", "Workbook 1 Value", "Workbook 2 Value", "Timestamp"};
    for (int column = 0; column < detailHeaders.size(); ++column)
        document.write(1, column + 1, detailHeaders.at(column), headerFormat());
    row = 2;
    for (const Change& change : result.changes) {
        if (change.type != ChangeType::Modified)
            continue;
        const QXlsx::Format detailFormat = statusFormat(change.type);
        for (const CellChange& cell : change.cells) {
            document.write(row, 1, change.sheet, detailFormat);
            document.write(row, 2, typeName(change.type), detailFormat);
            document.write(row, 3, change.key, detailFormat);
            document.write(row, 4, change.workbook1Key, detailFormat);
            document.write(row, 5, change.workbook2Key, detailFormat);
            document.write(row, 6, cell.column, detailFormat);
            writeExcelValue(document, row, 7, cell.column, cell.oldValue, detailFormat);
            writeExcelValue(document, row, 8, cell.column, cell.newValue, detailFormat);
            document.write(row, 9, result.created.toString(Qt::ISODate), detailFormat);
            ++row;
        }
    }
    document.setColumnWidth(1, 2, 20);
    document.setColumnWidth(3, 6, 26);
    document.setColumnWidth(7, 8, 38);
    document.setColumnWidth(9, 9, 24);
    document.selectSheet("Comparison Results");
    if (!document.saveAs(path))
        throw std::runtime_error("Cannot save Excel report.");
}

void ExcelEngine::exportMergedXlsx(const MergedSheet& sheet, const QString& path)
{
    QXlsx::Document document;
    document.renameSheet("Sheet1", sheet.name.isEmpty() ? "Merged Data" : sheet.name);
    for (int column = 0; column < sheet.headers.size(); ++column)
        document.write(1, column + 1, sheet.headers.at(column), headerFormat());

    for (int row = 0; row < sheet.rows.size(); ++row) {
        const QStringList& values = sheet.rows.at(row);
        for (int column = 0; column < values.size(); ++column) {
            const QString header = column < sheet.headers.size() ? sheet.headers.at(column) : QString();
            writeExcelValue(document, row + 2, column + 1, header, values.at(column), QXlsx::Format(), column == 0);
        }
    }

    for (int column = 1; column <= sheet.headers.size(); ++column)
        document.setColumnWidth(column, column, column == 1 ? 20 : 28);

    if (!document.saveAs(path))
        throw std::runtime_error("Cannot save the merged Excel workbook.");
}

void ExcelEngine::exportPdf(const Result& result, const QString& path)
{
    QPdfWriter pdf(path);
    pdf.setPageSize(QPageSize(QPageSize::A4));
    pdf.setPageOrientation(QPageLayout::Landscape);
    pdf.setResolution(150);
    pdf.setTitle("Excel Compare Professional Report");
    pdf.setCreator("AAT-Tech Ltd");

    QPainter painter(&pdf);
    if (!painter.isActive())
        throw std::runtime_error("Cannot create PDF.");

    const int margin = 90;
    int y = margin;
    const int pageWidth = pdf.width();
    const int pageHeight = pdf.height();
    auto newPage = [&]() { pdf.newPage(); y = margin; };

    QFont titleFont("Arial", 18, QFont::Bold);
    QFont headingFont("Arial", 10, QFont::Bold);
    QFont bodyFont("Arial", 8);

    painter.setFont(titleFont);
    painter.drawText(margin, y, "Excel Compare Professional");
    y += 42;
    painter.setFont(headingFont);
    painter.drawText(margin, y, "AAT-Tech Ltd");
    y += 35;
    painter.setFont(bodyFont);
    painter.drawText(margin, y, "Workbook 1: " + result.oldFile); y += 24;
    painter.drawText(margin, y, "Workbook 2: " + result.newFile); y += 24;
    painter.drawText(margin, y, "Generated: " + result.created.toString("dd MMM yyyy HH:mm:ss")); y += 35;
    painter.setFont(headingFont);
    painter.drawText(margin, y,
        QString("New Records: %1    Unmatched: %2    Modified: %3    Matched: %4")
            .arg(result.added).arg(result.deleted).arg(result.modified).arg(result.unchanged));
    y += 45;

    for (const Change& change : result.changes) {
        if (y > pageHeight - 120)
            newPage();
        painter.setFont(headingFont);
        painter.drawText(margin, y, change.sheet + "  |  " + typeName(change.type) + "  |  " + change.key);
        y += 22;
        painter.setFont(bodyFont);
        painter.drawText(margin, y, QString("Workbook 1 key: %1    Workbook 2 key: %2")
                                      .arg(change.workbook1Key, change.workbook2Key));
        y += 20;
        painter.setFont(bodyFont);
        QStringList details;
        if (change.type == ChangeType::Modified) {
            for (const CellChange& cell : change.cells)
                details.append(QString("%1: %2 -> %3").arg(cell.column, cell.oldValue, cell.newValue));
        } else {
            details.append(typeName(change.type) + " record");
        }
        painter.drawText(QRect(margin, y, pageWidth - 2 * margin, 55), Qt::TextWordWrap, details.join("; "));
        y += 65;
    }
    painter.end();
}

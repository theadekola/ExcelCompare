#pragma once

#include "ExcelEngine.h"
#include <QMainWindow>

class QCheckBox; class QComboBox; class QLabel; class QLineEdit; class QListWidget;
class QPushButton; class QSpinBox; class QStackedWidget; class QTableWidget; class QTextBrowser; class QTextEdit;

class MainWindow : public QMainWindow
{
    Q_OBJECT
public:
    MainWindow();

private slots:
    void loadWorkbooks();
    void refreshWorkspace();
    void autoMapColumns();
    void selectKeyColumns();
    void compareWorkbooks();
    void previewMerge();
    void mergeWorkbooks();
    void exportExcel();
    void exportPdf();
    void exportAudit();
    void filterResults();
    void showDetails(int row, int column);
    void showPage(int row);

private:
    QWidget* buildComparePage();
    QWidget* buildResultsPage();
    QWidget* buildReportsPage();
    QWidget* buildAuditPage();
    QWidget* buildSettingsPage();
    QWidget* buildAboutPage();
    void applyStyle();
    void setBusy(bool busy, const QString& message = {});
    void loadSelectedWorkbook(const QString& path, bool workbook1);
    void updateLoadedState();
    void fillPreview(QTableWidget* table, const Sheet& sheet);
    void fillMergedPreview();
    Options mergeOptions() const;
    void addMapping(const QString& oldColumn, const QString& newColumn,
                    const QString& outputColumn, bool key, bool compare, bool include);
    void populateMappingRows(bool applyAutoSelections, bool selectInitialKey);
    void updateMappingOutput(int row);
    Options mappingOptions() const;
    void showResult();
    Result filteredResult() const;
    void refreshReportPreview();
    void addAudit(const QString& event, const QString& details);
    QString friendlyException() const;

    QListWidget* navigation = nullptr;
    QStackedWidget* pages = nullptr;

    QLineEdit* oldFileEdit = nullptr;
    QLineEdit* newFileEdit = nullptr;
    QComboBox* oldSheetCombo = nullptr;
    QComboBox* newSheetCombo = nullptr;
    QTableWidget* oldPreview = nullptr;
    QTableWidget* newPreview = nullptr;
    QTableWidget* mappingTable = nullptr;
    QCheckBox* caseSensitive = nullptr;
    QCheckBox* trimSpaces = nullptr;
    QCheckBox* includeUnchanged = nullptr;
    QCheckBox* autoMapCompare = nullptr;
    QCheckBox* autoMapMerge = nullptr;
    QCheckBox* enableMerge = nullptr;
    QComboBox* joinMode = nullptr;
    QComboBox* conflictMode = nullptr;
    QPushButton* compareButton = nullptr;
    QPushButton* previewMergeButton = nullptr;
    QPushButton* mergeButton = nullptr;
    QSpinBox* mergeColumnCount = nullptr;
    QTableWidget* mergedPreview = nullptr;
    QLabel* mergedPreviewInfo = nullptr;

    QLabel* addedLabel = nullptr;
    QLabel* deletedLabel = nullptr;
    QLabel* modifiedLabel = nullptr;
    QLabel* unchangedLabel = nullptr;
    QLabel* totalLabel = nullptr;
    QLineEdit* resultSearch = nullptr;
    QComboBox* statusFilter = nullptr;
    QTableWidget* resultTable = nullptr;
    QTextEdit* detailView = nullptr;

    QTextBrowser* reportPreview = nullptr;
    QTableWidget* auditTable = nullptr;
    QLabel* statusLabel = nullptr;

    Workbook oldWorkbook;
    Workbook newWorkbook;
    Result result;
    MergedSheet mergedSheet;
    bool loaded = false;
};

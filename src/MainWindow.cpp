#include "MainWindow.h"

#include <QAbstractItemView>
#include <QAction>
#include <QCheckBox>
#include <QButtonGroup>
#include <QColor>
#include <QComboBox>
#include <QDateTime>
#include <QDialog>
#include <QDialogButtonBox>
#include <QFileDialog>
#include <QFileInfo>
#include <QFutureWatcher>
#include <QGridLayout>
#include <QGroupBox>
#include <QHeaderView>
#include <QHBoxLayout>
#include <QIcon>
#include <QLabel>
#include <QLineEdit>
#include <QListWidget>
#include <QMessageBox>
#include <QPixmap>
#include <QRegularExpression>
#include <QScrollArea>
#include <QFrame>
#include <QSizePolicy>
#include <QPushButton>
#include <QSettings>
#include <QSet>
#include <QSplitter>
#include <QSpinBox>
#include <QStackedWidget>
#include <QStatusBar>
#include <QTableWidget>
#include <QTextBrowser>
#include <QTextEdit>
#include <QVBoxLayout>
#include <QtConcurrent>

namespace {
const QString NoColumn = "(none)";
QString fileFilter(){return "Spreadsheet Files (*.xlsx *.xls *.xlsb *.xlsm *.xlt *.xltx *.xltm *.ods *.fods *.csv *.tsv);;All Files (*.*)";}
QComboBox* mappingCombo(const QStringList& headers,const QString& selected){auto*c=new QComboBox;c->addItem(NoColumn,QString());for(const auto&h:headers)c->addItem(h,h);int i=c->findData(selected);c->setCurrentIndex(i>=0?i:0);return c;}
QString mapped(QTableWidget*t,int r,int c){auto*box=qobject_cast<QComboBox*>(t->cellWidget(r,c));return box?box->currentData().toString():QString();}
bool checked(QTableWidget*t,int r,int c){auto*i=t->item(r,c);return i&&i->checkState()==Qt::Checked;}
QTableWidgetItem* checkItem(bool on){auto*i=new QTableWidgetItem;i->setFlags(Qt::ItemIsEnabled|Qt::ItemIsSelectable|Qt::ItemIsUserCheckable);i->setCheckState(on?Qt::Checked:Qt::Unchecked);return i;}

struct LoadOutcome {
    bool ok = false;
    Workbook oldWorkbook;
    Workbook newWorkbook;
    QString error;
};

struct SingleLoadOutcome {
    bool ok = false;
    Workbook workbook;
    QString error;
};

struct CompareOutcome {
    bool ok = false;
    Result result;
    QString error;
};
}

namespace {
QString comparisonDisplayValue(const Change& change, int field)
{
    const QString oldValue = field < change.oldRow.size() ? change.oldRow.at(field) : QString();
    const QString newValue = field < change.newRow.size() ? change.newRow.at(field) : QString();
    if (change.type == ChangeType::Deleted) return oldValue;
    if (change.type == ChangeType::Modified && oldValue != newValue) return QString("%1 -> %2").arg(oldValue, newValue);
    return newValue.isEmpty() ? oldValue : newValue;
}
}

MainWindow::MainWindow(){
    setWindowTitle("Excel Compare Professional - AAT-Tech Ltd");
    setWindowIcon(QIcon(":/branding/assets/app-icon.png"));
    resize(1400,850);
    setMinimumSize(900,650);

    auto* central=new QWidget;auto* shell=new QHBoxLayout(central);shell->setContentsMargins(0,0,0,0);shell->setSpacing(0);
    auto* sidebar=new QWidget;sidebar->setObjectName("sidebar");sidebar->setFixedWidth(230);auto*side=new QVBoxLayout(sidebar);side->setContentsMargins(16,18,16,16);
    auto* brand=new QLabel("<b style='font-size:19px'>Excel Compare</b><br><span style='color:#cbd5e1'>Professional Edition</span><br><small>v2.0.3 · AAT-Tech Ltd</small>");brand->setObjectName("brand");brand->setTextFormat(Qt::RichText);side->addWidget(brand);
    navigation=new QListWidget;navigation->setObjectName("navigation");navigation->addItems({"▣  Compare Files","▤  Results","▧  Reports","▥  Audit Log","⚙  Settings","ⓘ  About"});navigation->setCurrentRow(0);side->addWidget(navigation,1);
    auto*ready=new QLabel("●  Ready");ready->setObjectName("sideReady");side->addWidget(ready);shell->addWidget(sidebar);

    pages=new QStackedWidget;pages->addWidget(buildComparePage());pages->addWidget(buildResultsPage());pages->addWidget(buildReportsPage());pages->addWidget(buildAuditPage());pages->addWidget(buildSettingsPage());pages->addWidget(buildAboutPage());shell->addWidget(pages,1);
    connect(navigation,&QListWidget::currentRowChanged,this,&MainWindow::showPage);
    statusLabel=new QLabel("Ready");statusBar()->addWidget(statusLabel,1);setCentralWidget(central);applyStyle();
}

QWidget* MainWindow::buildComparePage(){
    auto* page = new QWidget;
    auto* pageLayout = new QVBoxLayout(page);
    pageLayout->setContentsMargins(0,0,0,0);

    auto* scroll = new QScrollArea;
    scroll->setWidgetResizable(true);
    scroll->setFrameShape(QFrame::NoFrame);
    scroll->setHorizontalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    scroll->setVerticalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    pageLayout->addWidget(scroll);

    auto* content = new QWidget;
    content->setMinimumWidth(760);
    auto* root = new QVBoxLayout(content);
    root->setContentsMargins(22,18,22,18);
    root->setSpacing(12);
    scroll->setWidget(content);

    auto* title = new QLabel("Compare Files");
    title->setObjectName("pageTitle");
    root->addWidget(title);
    root->addWidget(new QLabel("Open two spreadsheets, preview them side by side, choose columns, then compare or merge."));

    auto* previewSplit = new QSplitter(Qt::Horizontal);
    previewSplit->setChildrenCollapsible(false);
    previewSplit->setHandleWidth(6);
    auto buildWorkbookPanel = [&](const QString& titleText, QLineEdit*& fileEdit,
                                  QComboBox*& sheetCombo, QTableWidget*& preview,
                                  QPushButton*& browseButton) {
        auto* box = new QGroupBox(titleText);
        box->setMinimumWidth(340);
        auto* layout = new QVBoxLayout(box);

        preview = new QTableWidget;
        preview->setMinimumHeight(210);
        preview->setSizePolicy(QSizePolicy::Expanding,QSizePolicy::Expanding);
        preview->setEditTriggers(QAbstractItemView::NoEditTriggers);
        preview->setAlternatingRowColors(true);
        preview->setSelectionBehavior(QAbstractItemView::SelectRows);
        preview->horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
        layout->addWidget(preview, 1);

        auto* workbookControls = new QHBoxLayout;
        workbookControls->setSpacing(8);
        workbookControls->addWidget(new QLabel("Worksheet"));
        sheetCombo = new QComboBox;
        sheetCombo->setEnabled(false);
        sheetCombo->setMinimumWidth(120);
        workbookControls->addWidget(sheetCombo, 0);

        fileEdit = new QLineEdit;
        fileEdit->setPlaceholderText("Select spreadsheet file...");
        workbookControls->addWidget(fileEdit, 1);

        browseButton = new QPushButton("Browse");
        workbookControls->addWidget(browseButton, 0);
        layout->addLayout(workbookControls);

        previewSplit->addWidget(box);
    };

    QPushButton* browse1 = nullptr;
    QPushButton* browse2 = nullptr;
    buildWorkbookPanel("Workbook 1", oldFileEdit, oldSheetCombo, oldPreview, browse1);
    buildWorkbookPanel("Workbook 2", newFileEdit, newSheetCombo, newPreview, browse2);
    previewSplit->setStretchFactor(0,1);
    previewSplit->setStretchFactor(1,1);
    previewSplit->setSizes({700,700});
    root->addWidget(previewSplit, 3);

    connect(browse1,&QPushButton::clicked,this,[this]{
        const auto path=QFileDialog::getOpenFileName(this,"Workbook 1",{},fileFilter());
        if(path.isEmpty()) return;
        oldFileEdit->setText(path);
        loadSelectedWorkbook(path,true);
    });
    connect(browse2,&QPushButton::clicked,this,[this]{
        const auto path=QFileDialog::getOpenFileName(this,"Workbook 2",{},fileFilter());
        if(path.isEmpty()) return;
        newFileEdit->setText(path);
        loadSelectedWorkbook(path,false);
    });
    connect(oldSheetCombo,&QComboBox::currentTextChanged,this,&MainWindow::refreshWorkspace);
    connect(newSheetCombo,&QComboBox::currentTextChanged,this,&MainWindow::refreshWorkspace);

    auto* mapBox = new QGroupBox("Column mapping and actions");
    auto* mapLayout = new QVBoxLayout(mapBox);
    mapLayout->setSpacing(10);

    auto* mapTools = new QHBoxLayout;
    auto* autoMap = new QPushButton("Auto-map columns");
    auto* keys = new QPushButton("Select key columns");
    autoMapCompare = new QCheckBox("Tick Compare");
    autoMapMerge = new QCheckBox("Tick Merge");
    autoMapCompare->setChecked(true);
    autoMapMerge->setChecked(false);
    autoMapCompare->setToolTip("When Auto-map columns is clicked, select mapped columns for comparison.");
    autoMapMerge->setToolTip("When Auto-map columns is clicked, select mapped columns for the merged output.");
    mapTools->addWidget(autoMap);
    mapTools->addWidget(autoMapCompare);
    mapTools->addWidget(autoMapMerge);
    mapTools->addSpacing(12);
    mapTools->addWidget(keys);
    mapTools->addStretch();
    mapLayout->addLayout(mapTools);

    mappingTable = new QTableWidget(0,6);
    mappingTable->setHorizontalHeaderLabels({"Merge","Workbook 1 column","Workbook 2 column","Key","Compare","Output column"});
    mappingTable->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    mappingTable->verticalHeader()->hide();
    mappingTable->setMinimumHeight(180);
    mappingTable->setSizePolicy(QSizePolicy::Expanding,QSizePolicy::Expanding);
    mapLayout->addWidget(mappingTable,1);

    auto* optionsRow = new QHBoxLayout;
    caseSensitive = new QCheckBox("Case sensitive");
    trimSpaces = new QCheckBox("Ignore leading/trailing spaces");
    trimSpaces->setChecked(true);
    includeUnchanged = new QCheckBox("Include unchanged records");
    includeUnchanged->setChecked(true);
    includeUnchanged->setToolTip("All comparison statuses are retained so the Results filter can show each category.");
    optionsRow->addWidget(caseSensitive);
    optionsRow->addWidget(trimSpaces);
    optionsRow->addWidget(includeUnchanged);
    optionsRow->addStretch();
    mapLayout->addLayout(optionsRow);

    auto* separator = new QFrame;
    separator->setFrameShape(QFrame::HLine);
    separator->setFrameShadow(QFrame::Sunken);
    mapLayout->addWidget(separator);

    enableMerge = new QCheckBox;
    enableMerge->setChecked(false);
    enableMerge->hide();

    auto* modeRow = new QHBoxLayout;
    auto* compareOnlyButton = new QPushButton("Compare only");
    auto* compareMergeButton = new QPushButton("Compare and merge");
    compareButton = new QPushButton("Start Comparison");
    compareButton->setObjectName("primaryButton");
    compareButton->setMinimumHeight(42);
    compareButton->setMinimumWidth(220);
    compareButton->setEnabled(false);
    compareOnlyButton->setCheckable(true);
    compareMergeButton->setCheckable(true);
    compareOnlyButton->setChecked(true);
    compareOnlyButton->setObjectName("modeButton");
    compareMergeButton->setObjectName("modeButton");
    auto* modeGroup = new QButtonGroup(page);
    modeGroup->setExclusive(true);
    modeGroup->addButton(compareOnlyButton,0);
    modeGroup->addButton(compareMergeButton,1);
    modeRow->addWidget(new QLabel("Operation"));
    modeRow->addWidget(compareOnlyButton);
    modeRow->addWidget(compareMergeButton);
    modeRow->addStretch(1);
    modeRow->addWidget(compareButton, 0, Qt::AlignRight);
    mapLayout->addLayout(modeRow);

    auto* mergeControls = new QWidget;
    auto* controls = new QGridLayout(mergeControls);
    controls->setContentsMargins(0,0,0,0);
    joinMode = new QComboBox;
    joinMode->addItem("Full merge: include records from both workbooks",int(MergeJoin::FullOuter));
    joinMode->addItem("Keep all Workbook 1 records",int(MergeJoin::LeftOnly));
    joinMode->addItem("Matched records only",int(MergeJoin::Inner));
    conflictMode = new QComboBox;
    conflictMode->addItem("Prefer Workbook 2",int(ConflictPolicy::PreferNew));
    conflictMode->addItem("Prefer Workbook 1",int(ConflictPolicy::PreferOld));
    conflictMode->addItem("Show both conflicting values",int(ConflictPolicy::MarkConflict));
    mergeColumnCount = new QSpinBox;
    mergeColumnCount->setMinimum(1);
    mergeColumnCount->setMaximum(1);
    mergeColumnCount->setValue(1);
    mergeColumnCount->setToolTip("Maximum number of checked mapping rows to include in the merged output.");
    controls->addWidget(new QLabel("Merge rows"),0,0);
    controls->addWidget(joinMode,0,1);
    controls->addWidget(new QLabel("Number of columns to merge"),0,2);
    controls->addWidget(mergeColumnCount,0,3);
    controls->addWidget(new QLabel("Conflicting values"),1,0);
    controls->addWidget(conflictMode,1,1,1,3);
    controls->setColumnStretch(1,2);
    controls->setColumnStretch(3,1);
    mergeControls->setVisible(false);
    mapLayout->addWidget(mergeControls);

    auto* mergeActionRow = new QHBoxLayout;
    previewMergeButton = new QPushButton("Preview merged output");
    previewMergeButton->setObjectName("previewButton");
    previewMergeButton->setMinimumHeight(42);
    previewMergeButton->setMinimumWidth(220);
    previewMergeButton->setVisible(false);
    previewMergeButton->setEnabled(false);
    mergeActionRow->addStretch(1);
    mergeActionRow->addWidget(previewMergeButton, 0, Qt::AlignRight);
    mapLayout->addLayout(mergeActionRow);

    mergedPreviewInfo = new QLabel("Create a merge preview before saving. The preview shows the first 100 output rows.");
    mergedPreviewInfo->setObjectName("hintLabel");
    mergedPreviewInfo->setVisible(false);
    mapLayout->addWidget(mergedPreviewInfo);

    mergedPreview = new QTableWidget;
    mergedPreview->setMinimumHeight(190);
    mergedPreview->setEditTriggers(QAbstractItemView::NoEditTriggers);
    mergedPreview->setAlternatingRowColors(true);
    mergedPreview->setSelectionBehavior(QAbstractItemView::SelectRows);
    mergedPreview->setVisible(false);
    mapLayout->addWidget(mergedPreview,1);

    mergeButton = new QPushButton("Save preview as Excel");
    mergeButton->setObjectName("mergeButton");
    mergeButton->setEnabled(false);
    mergeButton->setVisible(false);
    auto* saveRow = new QHBoxLayout;
    saveRow->addStretch();
    saveRow->addWidget(mergeButton);
    mapLayout->addLayout(saveRow);

    root->addWidget(mapBox,3);

    connect(autoMap,&QPushButton::clicked,this,&MainWindow::autoMapColumns);
    connect(keys,&QPushButton::clicked,this,&MainWindow::selectKeyColumns);
    connect(mappingTable,&QTableWidget::itemChanged,this,[this](QTableWidgetItem* item){
        if(item && (item->column()==0 || item->column()==4)) updateMappingOutput(item->row());
    });
    connect(compareButton,&QPushButton::clicked,this,&MainWindow::compareWorkbooks);
    connect(modeGroup,qOverload<int>(&QButtonGroup::idClicked),this,[this,mergeControls](int id){
        const bool enabled = id==1;
        enableMerge->setChecked(enabled);
        mergeControls->setVisible(enabled);
        previewMergeButton->setVisible(enabled);
        mergedPreviewInfo->setVisible(enabled);
        mergedPreview->setVisible(enabled);
        mergeButton->setVisible(enabled);
        joinMode->setEnabled(enabled);
        conflictMode->setEnabled(enabled);
        mergeColumnCount->setEnabled(enabled);
        previewMergeButton->setEnabled(enabled && loaded);
        if(!enabled){
            mergedSheet = {};
            mergedPreview->clear();
            mergedPreview->setRowCount(0);
            mergedPreview->setColumnCount(0);
            mergeButton->setEnabled(false);
        }
    });
    connect(previewMergeButton,&QPushButton::clicked,this,&MainWindow::previewMerge);
    connect(mergeButton,&QPushButton::clicked,this,&MainWindow::mergeWorkbooks);
    connect(mappingTable,&QTableWidget::itemChanged,this,[this](QTableWidgetItem*){
        mergedSheet = {};
        if(mergeButton) mergeButton->setEnabled(false);
    });
    connect(mergeColumnCount,qOverload<int>(&QSpinBox::valueChanged),this,[this](int){
        mergedSheet={};
        if(mergeButton)mergeButton->setEnabled(false);
    });
    connect(joinMode,qOverload<int>(&QComboBox::currentIndexChanged),this,[this](int){
        mergedSheet={};
        if(mergeButton)mergeButton->setEnabled(false);
    });
    connect(conflictMode,qOverload<int>(&QComboBox::currentIndexChanged),this,[this](int){
        mergedSheet={};
        if(mergeButton)mergeButton->setEnabled(false);
    });
    return page;
}
QWidget* MainWindow::buildResultsPage(){
    auto*page=new QWidget;auto*root=new QVBoxLayout(page);root->setContentsMargins(22,18,22,18);auto*title=new QLabel("Comparison Summary");title->setObjectName("pageTitle");root->addWidget(title);
    auto*cards=new QHBoxLayout;auto card=[&](const QString&n,QLabel*&v,const QString&o){auto*b=new QGroupBox(n);b->setObjectName(o);auto*l=new QVBoxLayout(b);v=new QLabel("0");v->setObjectName("number");l->addWidget(v);cards->addWidget(b);};card("➕ New Records",addedLabel,"addedCard");card("⚠ Unmatched",deletedLabel,"deletedCard");card("✏ Modified",modifiedLabel,"modifiedCard");card("✓ Matched",unchangedLabel,"unchangedCard");card("📊 Total Records",totalLabel,"totalCard");root->addLayout(cards);
    auto*filters=new QHBoxLayout;
    resultSearch=new QLineEdit;
    resultSearch->setPlaceholderText("Search results...");
    resultSearch->setMaximumWidth(520);
    statusFilter=new QComboBox;
    statusFilter->addItems({"📋 All Results","➕ New Records","⚠ Unmatched","✏ Modified","✓ Matched"});
    auto*exportXlsxButton=new QPushButton("Export Excel");
    auto*exportPdfButton=new QPushButton("Export PDF");
    auto*exportCsvButton=new QPushButton("Export CSV");
    filters->addWidget(resultSearch);
    filters->addWidget(statusFilter);
    filters->addSpacing(8);
    filters->addWidget(exportXlsxButton);
    filters->addWidget(exportPdfButton);
    filters->addWidget(exportCsvButton);
    filters->addStretch(1);
    root->addLayout(filters);
    connect(resultSearch,&QLineEdit::textChanged,this,&MainWindow::filterResults);
    connect(statusFilter,&QComboBox::currentTextChanged,this,&MainWindow::filterResults);
    connect(exportXlsxButton,&QPushButton::clicked,this,&MainWindow::exportExcel);
    connect(exportPdfButton,&QPushButton::clicked,this,&MainWindow::exportPdf);
    connect(exportCsvButton,&QPushButton::clicked,this,&MainWindow::exportAudit);
    resultTable=new QTableWidget(0,3);resultTable->setHorizontalHeaderLabels({"Worksheet pair","Status","Record key"});resultTable->horizontalHeader()->setStretchLastSection(false);resultTable->setEditTriggers(QAbstractItemView::NoEditTriggers);resultTable->setSelectionBehavior(QAbstractItemView::SelectRows);root->addWidget(resultTable,1);detailView=new QTextEdit;detailView->setReadOnly(true);detailView->setMaximumHeight(130);detailView->setPlaceholderText("Double-click a result row to view field-level changes.");root->addWidget(detailView);connect(resultTable,&QTableWidget::cellDoubleClicked,this,&MainWindow::showDetails);return page;
}

QWidget* MainWindow::buildReportsPage(){auto*p=new QWidget;auto*l=new QVBoxLayout(p);l->setContentsMargins(22,18,22,18);auto*t=new QLabel("Comparison Report");t->setObjectName("pageTitle");l->addWidget(t);auto*bar=new QHBoxLayout;auto*pdf=new QPushButton("Export PDF");auto*xlsx=new QPushButton("Export Excel");auto*csv=new QPushButton("Export Audit CSV");bar->addWidget(pdf);bar->addWidget(xlsx);bar->addWidget(csv);bar->addStretch();l->addLayout(bar);reportPreview=new QTextBrowser;l->addWidget(reportPreview,1);connect(pdf,&QPushButton::clicked,this,&MainWindow::exportPdf);connect(xlsx,&QPushButton::clicked,this,&MainWindow::exportExcel);connect(csv,&QPushButton::clicked,this,&MainWindow::exportAudit);refreshReportPreview();return p;}
QWidget* MainWindow::buildAuditPage(){auto*p=new QWidget;auto*l=new QVBoxLayout(p);l->setContentsMargins(22,18,22,18);auto*t=new QLabel("Audit Log");t->setObjectName("pageTitle");l->addWidget(t);auto*bar=new QHBoxLayout;auto*clear=new QPushButton("Clear log");auto*exportB=new QPushButton("Export CSV");bar->addWidget(exportB);bar->addWidget(clear);bar->addStretch();l->addLayout(bar);auditTable=new QTableWidget(0,3);auditTable->setHorizontalHeaderLabels({"Date / Time","Event","Details"});auditTable->horizontalHeader()->setSectionResizeMode(2,QHeaderView::Stretch);auditTable->setEditTriggers(QAbstractItemView::NoEditTriggers);l->addWidget(auditTable);connect(clear,&QPushButton::clicked,this,[this]{auditTable->setRowCount(0);});connect(exportB,&QPushButton::clicked,this,&MainWindow::exportAudit);return p;}
QWidget* MainWindow::buildSettingsPage(){auto*p=new QWidget;auto*l=new QVBoxLayout(p);l->setContentsMargins(22,18,22,18);auto*t=new QLabel("Settings");t->setObjectName("pageTitle");l->addWidget(t);auto*g=new QGroupBox("Default comparison preferences");auto*gl=new QVBoxLayout(g);auto*c1=new QCheckBox("Case-sensitive matching by default");auto*c2=new QCheckBox("Ignore leading and trailing spaces");c2->setChecked(true);auto*c3=new QCheckBox("Include unchanged records in reports");gl->addWidget(c1);gl->addWidget(c2);gl->addWidget(c3);l->addWidget(g);auto*info=new QLabel("Legacy .xls files open natively. LibreOffice is used only for .xlsb, .ods, .fods and other conversion-only formats.");info->setWordWrap(true);l->addWidget(info);l->addStretch();return p;}
QWidget* MainWindow::buildAboutPage(){auto*p=new QWidget;auto*l=new QVBoxLayout(p);l->setAlignment(Qt::AlignCenter);auto*icon=new QLabel;icon->setPixmap(QPixmap(":/branding/assets/app-icon.png").scaled(120,120,Qt::KeepAspectRatio,Qt::SmoothTransformation));icon->setAlignment(Qt::AlignCenter);l->addWidget(icon);auto*t=new QLabel("Excel Compare Professional");t->setObjectName("pageTitle");t->setAlignment(Qt::AlignCenter);l->addWidget(t);auto*d=new QLabel("Version 2.0.3<br><b>Developed and published by AAT-Tech Ltd</b><br><br>Professional spreadsheet comparison, reporting and merge software.<br><br>© 2026 AAT-Tech Ltd. All Rights Reserved.");d->setTextFormat(Qt::RichText);d->setAlignment(Qt::AlignCenter);l->addWidget(d);return p;}

void MainWindow::showPage(int row){if(row>=0&&row<pages->count())pages->setCurrentIndex(row);}
void MainWindow::setBusy(bool busy,const QString& message){
    if(compareButton)compareButton->setEnabled(!busy&&loaded);
    const bool mergeEnabled = enableMerge && enableMerge->isChecked();
    if(previewMergeButton)previewMergeButton->setEnabled(!busy&&loaded&&mergeEnabled);
    if(mergeButton)mergeButton->setEnabled(!busy&&loaded&&mergeEnabled&&!mergedSheet.headers.isEmpty());
    statusLabel->setText(message.isEmpty()?(busy?"Working...":"Ready"):message);
}
QString MainWindow::friendlyException() const{return "The workbook could not be loaded. For .xls/.xlsb/.ods files, confirm LibreOffice is installed and close the spreadsheet in Excel before trying again.";}

void MainWindow::loadSelectedWorkbook(const QString& path, bool workbook1){
    if(path.trimmed().isEmpty()) return;

    setBusy(true,QString("Loading Workbook %1...").arg(workbook1?1:2));
    addAudit("Workbook load started",QFileInfo(path).fileName());

    auto* watcher=new QFutureWatcher<SingleLoadOutcome>(this);
    connect(watcher,&QFutureWatcher<SingleLoadOutcome>::finished,this,[this,watcher,workbook1,path]{
        const SingleLoadOutcome outcome=watcher->result();
        if(!outcome.ok){
            setBusy(false,QString("Workbook %1 load failed.").arg(workbook1?1:2));
            addAudit("Workbook load failed",outcome.error);
            QMessageBox::critical(this,"Open spreadsheet",outcome.error);
            watcher->deleteLater();
            return;
        }

        QComboBox* combo=workbook1?oldSheetCombo:newSheetCombo;
        QTableWidget* preview=workbook1?oldPreview:newPreview;
        Workbook& target=workbook1?oldWorkbook:newWorkbook;
        target=outcome.workbook;

        combo->blockSignals(true);
        combo->clear();
        combo->addItems(target.sheets.keys());
        combo->setEnabled(combo->count()>0);
        combo->blockSignals(false);

        if(combo->count()>0 && target.sheets.contains(combo->currentText()))
            fillPreview(preview,target.sheets.value(combo->currentText()));
        else
            preview->clear();

        updateLoadedState();
        setBusy(false,QString("Workbook %1 loaded and previewed.").arg(workbook1?1:2));
        addAudit("Workbook loaded",QString("Workbook %1: %2 sheet(s)").arg(workbook1?1:2).arg(target.sheets.size()));
        watcher->deleteLater();
    });

    watcher->setFuture(QtConcurrent::run([path]{
        SingleLoadOutcome outcome;
        try{
            ExcelEngine engine;
            outcome.workbook=engine.read(path);
            outcome.ok=true;
        }catch(const std::exception& e){
            outcome.error=QString::fromUtf8(e.what());
        }catch(...){
            outcome.error="An unexpected error occurred while opening the spreadsheet.";
        }
        return outcome;
    }));
}

void MainWindow::updateLoadedState(){
    loaded=oldSheetCombo->count()>0 && newSheetCombo->count()>0
        && !oldWorkbook.sheets.isEmpty() && !newWorkbook.sheets.isEmpty();

    if(loaded){
        refreshWorkspace();
        if(statusLabel) statusLabel->setText("Both workbooks loaded. Select a key column and start comparison.");
    }else{
        if(mappingTable) mappingTable->setRowCount(0);
        if(compareButton) compareButton->setEnabled(false);
        if(previewMergeButton) previewMergeButton->setEnabled(false);
        if(mergeButton) mergeButton->setEnabled(false);
        if(statusLabel) statusLabel->setText("Select the remaining workbook. Each selected file is previewed automatically.");
    }
}

void MainWindow::loadWorkbooks(){
    const QString a=oldFileEdit->text().trimmed(),b=newFileEdit->text().trimmed();
    if(a.isEmpty()||b.isEmpty()){QMessageBox::warning(this,"Missing files","Select both spreadsheet files.");return;}
    loaded=false;
    setBusy(true,"Loading spreadsheets...");
    addAudit("Load started",QFileInfo(a).fileName()+" vs "+QFileInfo(b).fileName());

    auto*w=new QFutureWatcher<LoadOutcome>(this);
    connect(w,&QFutureWatcher<LoadOutcome>::finished,this,[this,w]{
        const LoadOutcome outcome=w->result();
        if(!outcome.ok){
            loaded=false;
            setBusy(false,"Load failed.");
            addAudit("Load failed",outcome.error);
            QMessageBox::critical(this,"Open spreadsheet",outcome.error);
            w->deleteLater();
            return;
        }

        oldWorkbook=outcome.oldWorkbook;
        newWorkbook=outcome.newWorkbook;
        oldSheetCombo->blockSignals(true);
        newSheetCombo->blockSignals(true);
        oldSheetCombo->clear();
        newSheetCombo->clear();
        oldSheetCombo->addItems(oldWorkbook.sheets.keys());
        newSheetCombo->addItems(newWorkbook.sheets.keys());
        oldSheetCombo->setEnabled(oldSheetCombo->count()>0);
        newSheetCombo->setEnabled(newSheetCombo->count()>0);
        oldSheetCombo->blockSignals(false);
        newSheetCombo->blockSignals(false);
        loaded=oldSheetCombo->count()>0&&newSheetCombo->count()>0;
        refreshWorkspace();
        setBusy(false,loaded?"Workbooks loaded successfully.":"No readable worksheets found.");
        addAudit("Workbooks loaded",QString("%1 sheet(s) and %2 sheet(s)").arg(oldWorkbook.sheets.size()).arg(newWorkbook.sheets.size()));
        w->deleteLater();
    });

    w->setFuture(QtConcurrent::run([a,b]{
        LoadOutcome outcome;
        try{
            ExcelEngine engine;
            outcome.oldWorkbook=engine.read(a);
            outcome.newWorkbook=engine.read(b);
            outcome.ok=true;
        }catch(const std::exception&e){
            outcome.error=QString::fromUtf8(e.what());
        }catch(...){
            outcome.error="An unexpected error occurred while opening the spreadsheets.";
        }
        return outcome;
    }));
}

void MainWindow::fillPreview(QTableWidget*t,const Sheet&s){t->clear();t->setColumnCount(s.headers.size());t->setHorizontalHeaderLabels(s.headers);int n=qMin(50,int(s.rows.size()));t->setRowCount(n);for(int r=0;r<n;++r)for(int c=0;c<s.headers.size();++c)t->setItem(r,c,new QTableWidgetItem(c<s.rows[r].size()?s.rows[r][c]:QString()));t->horizontalHeader()->setSectionResizeMode(QHeaderView::ResizeToContents);}
void MainWindow::refreshWorkspace(){
    if(!loaded||!oldWorkbook.sheets.contains(oldSheetCombo->currentText())||!newWorkbook.sheets.contains(newSheetCombo->currentText()))return;
    fillPreview(oldPreview,oldWorkbook.sheets.value(oldSheetCombo->currentText()));
    fillPreview(newPreview,newWorkbook.sheets.value(newSheetCombo->currentText()));

    // Populate the mapping grid immediately after loading or changing a worksheet.
    // The columns are paired by matching header names, but Compare and Merge remain
    // unticked until the user chooses them or presses Auto-map columns.
    populateMappingRows(false,true);

    mergedSheet={};
    if(mergeButton)mergeButton->setEnabled(false);
    if(compareButton)compareButton->setEnabled(loaded && mappingTable->rowCount()>0);
    if(previewMergeButton)previewMergeButton->setEnabled(false);
    if(statusLabel)statusLabel->setText("Workbooks loaded. Select a key column, then tick Compare and/or Merge for the required columns.");
}

void MainWindow::populateMappingRows(bool applyAutoSelections,bool selectInitialKey){
    if(!loaded)return;
    const auto& a=oldWorkbook.sheets.value(oldSheetCombo->currentText());
    const auto& b=newWorkbook.sheets.value(newSheetCombo->currentText());
    mappingTable->blockSignals(true);
    mappingTable->setRowCount(0);
    QSet<QString> used;
    bool keyAssigned=false;
    for(const auto& h:a.headers){
        QString match;
        for(const auto& n:b.headers){
            if(!used.contains(n)&&h.compare(n,Qt::CaseInsensitive)==0){match=n;break;}
        }
        if(!match.isEmpty())used.insert(match);
        const bool matched=!match.isEmpty();
        const bool compareSelected = applyAutoSelections
            ? (autoMapCompare && autoMapCompare->isChecked())
            : true;
        const bool mergeSelected = applyAutoSelections && matched && autoMapMerge && autoMapMerge->isChecked();
        const bool makeKey=selectInitialKey && matched && !keyAssigned;
        addMapping(h,match,(compareSelected||mergeSelected)?(match.isEmpty()?h:match):QString(),makeKey,compareSelected,mergeSelected);
        if(makeKey)keyAssigned=true;
    }
    for(const auto& h:b.headers){
        if(!used.contains(h)){
            const bool mergeSelected=applyAutoSelections && autoMapMerge && autoMapMerge->isChecked();
            addMapping({},h,mergeSelected?h:QString(),false,false,mergeSelected);
        }
    }
    mappingTable->blockSignals(false);
    if(mergeColumnCount){
        mergeColumnCount->setMaximum(qMax(1,mappingTable->rowCount()));
        int mergeCount=0;
        for(int r=0;r<mappingTable->rowCount();++r) if(checked(mappingTable,r,0)) ++mergeCount;
        mergeColumnCount->setValue(qMax(1,mergeCount));
    }
}

void MainWindow::autoMapColumns(){
    if(!loaded)return;
    populateMappingRows(true,false);
    mergedSheet={};
    if(mergeButton)mergeButton->setEnabled(false);
    if(compareButton)compareButton->setEnabled(loaded && mappingTable->rowCount()>0);
    if(previewMergeButton)previewMergeButton->setEnabled(loaded && enableMerge && enableMerge->isChecked() && mappingTable->rowCount()>0);
    setBusy(false,QString("Mapped %1 column rows. Compare ticks: %2. Merge ticks: %3. Select one or more key columns before starting.")
        .arg(mappingTable->rowCount())
        .arg(autoMapCompare&&autoMapCompare->isChecked()?"on":"off")
        .arg(autoMapMerge&&autoMapMerge->isChecked()?"on":"off"));
}

void MainWindow::updateMappingOutput(int row){
    if(!mappingTable || row<0 || row>=mappingTable->rowCount())return;
    const bool active=checked(mappingTable,row,0)||checked(mappingTable,row,4);
    QString output;
    if(active){
        const QString oldColumn=mapped(mappingTable,row,1);
        const QString newColumn=mapped(mappingTable,row,2);
        output=!newColumn.isEmpty()?newColumn:oldColumn;
    }
    mappingTable->blockSignals(true);
    if(!mappingTable->item(row,5)) mappingTable->setItem(row,5,new QTableWidgetItem);
    mappingTable->item(row,5)->setText(output);
    mappingTable->blockSignals(false);
}

void MainWindow::addMapping(const QString&a,const QString&b,const QString&o,bool key,bool comp,bool use){
    int r=mappingTable->rowCount();
    mappingTable->insertRow(r);
    mappingTable->setItem(r,0,checkItem(use));
    auto* oldBox=mappingCombo(oldWorkbook.sheets.value(oldSheetCombo->currentText()).headers,a);
    auto* newBox=mappingCombo(newWorkbook.sheets.value(newSheetCombo->currentText()).headers,b);
    mappingTable->setCellWidget(r,1,oldBox);
    mappingTable->setCellWidget(r,2,newBox);
    mappingTable->setItem(r,3,checkItem(key));
    mappingTable->setItem(r,4,checkItem(comp));
    mappingTable->setItem(r,5,new QTableWidgetItem(o));
    connect(oldBox,qOverload<int>(&QComboBox::currentIndexChanged),this,[this,r](int){updateMappingOutput(r);});
    connect(newBox,qOverload<int>(&QComboBox::currentIndexChanged),this,[this,r](int){updateMappingOutput(r);});
}
void MainWindow::selectKeyColumns(){
    if(!loaded)return;
    if(mappingTable->rowCount()==0){
        QMessageBox::information(this,"Select Key Columns","Load both workbooks first.");
        return;
    }
    QDialog d(this);
    d.setWindowTitle("Select Key Columns");
    auto*l=new QVBoxLayout(&d);
    l->addWidget(new QLabel("Select one or more columns that exist in both workbooks and uniquely identify a record."));
    auto*list=new QListWidget;
    QVector<int> sourceRows;
    for(int r=0;r<mappingTable->rowCount();++r){
        const QString oldColumn=mapped(mappingTable,r,1);
        const QString newColumn=mapped(mappingTable,r,2);
        if(oldColumn.isEmpty()||newColumn.isEmpty())continue;
        QString name=mappingTable->item(r,5)?mappingTable->item(r,5)->text().trimmed():QString();
        if(name.isEmpty())name=!newColumn.isEmpty()?newColumn:oldColumn;
        auto*i=new QListWidgetItem(name,list);
        i->setFlags(i->flags()|Qt::ItemIsUserCheckable);
        i->setCheckState(checked(mappingTable,r,3)?Qt::Checked:Qt::Unchecked);
        sourceRows.append(r);
    }
    if(sourceRows.isEmpty()){
        delete list;
        QMessageBox::warning(this,"Select Key Columns","No columns are currently mapped in both workbooks.");
        return;
    }
    l->addWidget(list);
    auto*buttons=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel);
    l->addWidget(buttons);
    connect(buttons,&QDialogButtonBox::accepted,&d,&QDialog::accept);
    connect(buttons,&QDialogButtonBox::rejected,&d,&QDialog::reject);
    if(d.exec()==QDialog::Accepted){
        for(int r=0;r<mappingTable->rowCount();++r)mappingTable->item(r,3)->setCheckState(Qt::Unchecked);
        for(int i=0;i<list->count();++i)mappingTable->item(sourceRows.at(i),3)->setCheckState(list->item(i)->checkState());
    }
}
Options MainWindow::mappingOptions()const{Options o;o.caseSensitive=caseSensitive->isChecked();o.trim=trimSpaces->isChecked();o.includeUnchanged=includeUnchanged->isChecked();for(int r=0;r<mappingTable->rowCount();++r){ColumnMapping m;m.include=checked(mappingTable,r,0);m.oldColumn=mapped(mappingTable,r,1);m.newColumn=mapped(mappingTable,r,2);m.key=checked(mappingTable,r,3);m.compare=checked(mappingTable,r,4);m.outputColumn=mappingTable->item(r,5)?mappingTable->item(r,5)->text():QString();o.mappings.append(m);}return o;}
Options MainWindow::mergeOptions() const{
    Options options=mappingOptions();
    const int limit=mergeColumnCount?mergeColumnCount->value():options.mappings.size();
    int included=0;
    for(auto& mapping:options.mappings){
        if(mapping.include){
            ++included;
            if(included>limit)mapping.include=false;
        }
    }
    return options;
}

void MainWindow::fillMergedPreview(){
    if(!mergedPreview)return;
    mergedPreview->clear();
    mergedPreview->setColumnCount(mergedSheet.headers.size());
    mergedPreview->setHorizontalHeaderLabels(mergedSheet.headers);
    const int rows=qMin(100,int(mergedSheet.rows.size()));
    mergedPreview->setRowCount(rows);
    for(int r=0;r<rows;++r){
        const auto& values=mergedSheet.rows.at(r);
        for(int c=0;c<mergedSheet.headers.size();++c){
            mergedPreview->setItem(r,c,new QTableWidgetItem(c<values.size()?values.at(c):QString()));
        }
    }
    mergedPreview->horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    mergedPreviewInfo->setText(QString("Previewing %1 of %2 merged rows with %3 output columns.")
        .arg(rows).arg(mergedSheet.rows.size()).arg(mergedSheet.headers.size()));
}

void MainWindow::previewMerge(){
    if(!loaded)return;
    if(!enableMerge || !enableMerge->isChecked()){
        QMessageBox::information(this,"Compare only","Tick 'Enable merge (compare and merge)' to preview a merged output.");
        return;
    }
    try{
        ExcelEngine engine;
        mergedSheet=engine.mergeMapped(oldWorkbook,oldSheetCombo->currentText(),newWorkbook,
            newSheetCombo->currentText(),mergeOptions(),MergeJoin(joinMode->currentData().toInt()),
            ConflictPolicy(conflictMode->currentData().toInt()));
        if(mergedSheet.headers.isEmpty()){
            throw std::runtime_error("No merge columns were selected. Check the Merge boxes in Column mapping.");
        }
        fillMergedPreview();
        mergeButton->setEnabled(true);
        addAudit("Merge preview created",QString("%1 rows, %2 columns").arg(mergedSheet.rows.size()).arg(mergedSheet.headers.size()));
    }catch(const std::exception& error){
        mergedSheet={};
        mergeButton->setEnabled(false);
        addAudit("Merge preview failed",error.what());
        QMessageBox::critical(this,"Merge preview",error.what());
    }
}

void MainWindow::compareWorkbooks(){
    if(!loaded)return;
    auto o=mappingOptions();
    auto a=oldWorkbook,b=newWorkbook;
    auto as=oldSheetCombo->currentText(),bs=newSheetCombo->currentText();
    setBusy(true,"Comparing records...");
    addAudit("Comparison started",as+" vs "+bs);

    auto*w=new QFutureWatcher<CompareOutcome>(this);
    connect(w,&QFutureWatcher<CompareOutcome>::finished,this,[this,w]{
        const CompareOutcome outcome=w->result();
        if(!outcome.ok){
            setBusy(false,"Comparison failed.");
            addAudit("Comparison failed",outcome.error);
            QMessageBox::critical(this,"Comparison",outcome.error);
            w->deleteLater();
            return;
        }
        result=outcome.result;
        const qsizetype comparisonOutcomeCount = result.added + result.deleted + result.modified + result.unchanged;
        result.totalRecords = (enableMerge && enableMerge->isChecked())
            ? comparisonOutcomeCount
            : result.workbook1Records;
        showResult();
        refreshReportPreview();
        setBusy(false,QString("Comparison complete: %1 records").arg(result.added+result.deleted+result.modified+result.unchanged));
        addAudit("Comparison completed",QString("New Records %1, Unmatched %2, Modified %3, Matched %4").arg(result.added).arg(result.deleted).arg(result.modified).arg(result.unchanged));
        navigation->setCurrentRow(1);
        w->deleteLater();
    });

    w->setFuture(QtConcurrent::run([a,b,as,bs,o]{
        CompareOutcome outcome;
        try{
            ExcelEngine engine;
            outcome.result=engine.compareMapped(a,as,b,bs,o);
            outcome.ok=true;
        }catch(const std::exception&e){
            outcome.error=QString::fromUtf8(e.what());
        }catch(...){
            outcome.error="An unexpected error occurred during comparison.";
        }
        return outcome;
    }));
}

void MainWindow::mergeWorkbooks(){
    if(!enableMerge || !enableMerge->isChecked()){
        QMessageBox::information(this,"Compare only","Tick 'Enable merge (compare and merge)' before saving a merged workbook.");
        return;
    }
    if(mergedSheet.headers.isEmpty()){
        QMessageBox::information(this,"Merge","Preview the merged output before saving.");
        return;
    }
    const auto path=QFileDialog::getSaveFileName(this,"Save merged workbook","merged-workbook.xlsx","Excel Workbook (*.xlsx)");
    if(path.isEmpty())return;
    try{
        ExcelEngine::exportMergedXlsx(mergedSheet,path);
        addAudit("Merged workbook created",path);
        QMessageBox::information(this,"Merge complete",QString("Saved %1 rows and %2 columns.\n%3")
            .arg(mergedSheet.rows.size()).arg(mergedSheet.headers.size()).arg(path));
    }catch(const std::exception& error){
        addAudit("Merge save failed",error.what());
        QMessageBox::critical(this,"Merge",error.what());
    }
}
void MainWindow::showResult()
{
    addedLabel->setText(QString::number(result.added));
    deletedLabel->setText(QString::number(result.deleted));
    modifiedLabel->setText(QString::number(result.modified));
    unchangedLabel->setText(QString::number(result.unchanged));
    totalLabel->setText(QString::number(result.totalRecords));

    QStringList headers = {"Worksheet pair", "Status", "Record key", "Workbook 1 key", "Workbook 2 key"};
    headers.append(result.outputHeaders);
    headers.append("Changed columns");
    resultTable->clear();
    resultTable->setColumnCount(headers.size());
    resultTable->setHorizontalHeaderLabels(headers);
    resultTable->setRowCount(result.changes.size());

    for (int row = 0; row < result.changes.size(); ++row) {
        const Change& change = result.changes.at(row);
        const QColor colour = change.type == ChangeType::Added ? QColor("#dcfce7")
            : change.type == ChangeType::Deleted ? QColor("#fee2e2")
            : change.type == ChangeType::Modified ? QColor("#fef3c7")
            : QColor("#f1f5f9");

        QStringList values = {change.sheet, typeName(change.type), change.key, change.workbook1Key, change.workbook2Key};
        for (int field = 0; field < result.outputHeaders.size(); ++field) {
            values.append(comparisonDisplayValue(change, field));
        }

        QStringList changedColumns;
        for (const CellChange& cell : change.cells)
            changedColumns.append(cell.column);
        values.append(changedColumns.join(", "));

        for (int column = 0; column < values.size(); ++column) {
            auto* item = new QTableWidgetItem(values.at(column));
            item->setBackground(colour);
            resultTable->setItem(row, column, item);
        }
    }

    resultTable->horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    resultTable->horizontalHeader()->setSectionResizeMode(0, QHeaderView::ResizeToContents);
    resultTable->horizontalHeader()->setSectionResizeMode(1, QHeaderView::ResizeToContents);
    resultTable->horizontalHeader()->setSectionResizeMode(2, QHeaderView::ResizeToContents);
    resultTable->horizontalHeader()->setSectionResizeMode(3, QHeaderView::ResizeToContents);
    resultTable->horizontalHeader()->setSectionResizeMode(4, QHeaderView::ResizeToContents);
    resultTable->resizeColumnsToContents();
    for (int column = 5; column < resultTable->columnCount(); ++column) {
        if (resultTable->columnWidth(column) > 260)
            resultTable->setColumnWidth(column, 260);
        else if (resultTable->columnWidth(column) < 110)
            resultTable->setColumnWidth(column, 110);
    }
    filterResults();
}
void MainWindow::filterResults()
{
    if (!resultTable)
        return;

    const QString query = resultSearch ? resultSearch->text().trimmed() : QString();
    const QString selection = statusFilter ? statusFilter->currentText() : QString("📋 All Results");

    QString expectedStatus;
    if (selection.contains("New Record", Qt::CaseInsensitive))
        expectedStatus = "➕ New Record";
    else if (selection.contains("Unmatched", Qt::CaseInsensitive))
        expectedStatus = "⚠ Unmatched";
    else if (selection.contains("Modified", Qt::CaseInsensitive))
        expectedStatus = "✏ Modified";
    else if (selection.contains("Matched", Qt::CaseInsensitive))
        expectedStatus = "✓ Matched";

    const bool showAll = selection.contains("All Results", Qt::CaseInsensitive);

    for (int row = 0; row < resultTable->rowCount(); ++row) {
        const QTableWidgetItem* statusItem = resultTable->item(row, 1);
        const bool statusMatches = showAll ||
            (statusItem && statusItem->text().compare(expectedStatus, Qt::CaseInsensitive) == 0);

        bool textMatches = query.isEmpty();
        for (int column = 0; column < resultTable->columnCount() && !textMatches; ++column) {
            const QTableWidgetItem* item = resultTable->item(row, column);
            if (item && item->text().contains(query, Qt::CaseInsensitive))
                textMatches = true;
        }

        resultTable->setRowHidden(row, !(statusMatches && textMatches));
    }
}
Result MainWindow::filteredResult() const
{
    Result output = result;
    output.changes.clear();
    output.added = 0;
    output.deleted = 0;
    output.modified = 0;
    output.unchanged = 0;

    if (!resultTable)
        return output;

    const int rows = qMin(resultTable->rowCount(), static_cast<int>(result.changes.size()));
    for (int row = 0; row < rows; ++row) {
        if (resultTable->isRowHidden(row))
            continue;

        const Change& change = result.changes.at(row);
        output.changes.append(change);
        switch (change.type) {
        case ChangeType::Added: ++output.added; break;
        case ChangeType::Deleted: ++output.deleted; break;
        case ChangeType::Modified: ++output.modified; break;
        case ChangeType::Unchanged: ++output.unchanged; break;
        }
    }
    output.totalRecords = output.changes.size();
    return output;
}
void MainWindow::showDetails(int row, int)
{
    if (row < 0 || row >= result.changes.size())
        return;
    const Change& change = result.changes.at(row);
    QString html = QString("<h3>%1</h3><b>Worksheet:</b> %2<br><b>Record key:</b> %3<br><b>Workbook 1 key:</b> %4<br><b>Workbook 2 key:</b> %5<hr>")
        .arg(typeName(change.type), change.sheet.toHtmlEscaped(), change.key.toHtmlEscaped(),
             change.workbook1Key.toHtmlEscaped(), change.workbook2Key.toHtmlEscaped());
    html += "<table cellspacing='0' cellpadding='6' border='1'><tr><th>Column</th><th>Workbook 1</th><th>Workbook 2</th></tr>";
    for (int field = 0; field < result.outputHeaders.size(); ++field) {
        const QString oldValue = field < change.oldRow.size() ? change.oldRow.at(field) : QString();
        const QString newValue = field < change.newRow.size() ? change.newRow.at(field) : QString();
        const bool changed = oldValue != newValue;
        html += QString("<tr%1><td><b>%2</b></td><td>%3</td><td>%4</td></tr>")
            .arg(changed ? " style='background:#fef3c7'" : "",
                 result.outputHeaders.value(field).toHtmlEscaped(),
                 oldValue.toHtmlEscaped(), newValue.toHtmlEscaped());
    }
    html += "</table>";
    detailView->setHtml(html);
}
void MainWindow::refreshReportPreview(){if(!reportPreview)return;reportPreview->setHtml(QString("<h1 style='color:#172554'>Excel Compare Professional Report</h1><p><b>AAT-Tech Ltd</b></p><p>Workbook 1: %1<br>Workbook 2: %2<br>Generated: %3</p><table cellpadding='10' border='1' cellspacing='0'><tr><th>➕ New Records</th><th>⚠ Unmatched</th><th>✏ Modified</th><th>✓ Matched</th><th>📊 Total Records</th></tr><tr><td>%4</td><td>%5</td><td>%6</td><td>%7</td><td><b>%8</b></td></tr></table>").arg(result.oldFile.toHtmlEscaped(),result.newFile.toHtmlEscaped(),result.created.isValid()?result.created.toString("dd MMM yyyy HH:mm:ss"):"No comparison completed").arg(result.added).arg(result.deleted).arg(result.modified).arg(result.unchanged).arg(result.totalRecords));}
void MainWindow::addAudit(const QString&e,const QString&d){if(!auditTable)return;int r=auditTable->rowCount();auditTable->insertRow(r);auditTable->setItem(r,0,new QTableWidgetItem(QDateTime::currentDateTime().toString("dd/MM/yyyy HH:mm:ss")));auditTable->setItem(r,1,new QTableWidgetItem(e));auditTable->setItem(r,2,new QTableWidgetItem(d));}
void MainWindow::exportExcel()
{
    if (result.changes.isEmpty()) {
        QMessageBox::information(this, "Export", "Run a comparison first.");
        return;
    }
    const Result visible = filteredResult();
    if (visible.changes.isEmpty()) {
        QMessageBox::information(this, "Export", "No records match the selected status and search filter.");
        return;
    }
    const QString selection = statusFilter ? statusFilter->currentText() : QString("📋 All Results");
    QString safeSelection = selection;
    safeSelection.remove(QRegularExpression("[^A-Za-z0-9 -]"));
    safeSelection = safeSelection.trimmed().toLower().replace(' ', '-');
    const QString defaultName = QString("comparison-%1.xlsx").arg(safeSelection);
    const QString path = QFileDialog::getSaveFileName(this, "Export visible results to Excel", defaultName, "Excel (*.xlsx)");
    if (path.isEmpty()) return;
    try {
        ExcelEngine::exportXlsx(visible, path);
        addAudit("Excel results exported", QString("%1: %2 records").arg(selection).arg(visible.changes.size()));
    } catch (const std::exception& error) {
        QMessageBox::critical(this, "Export", error.what());
    }
}

void MainWindow::exportPdf()
{
    if (result.changes.isEmpty()) {
        QMessageBox::information(this, "Export", "Run a comparison first.");
        return;
    }
    const Result visible = filteredResult();
    if (visible.changes.isEmpty()) {
        QMessageBox::information(this, "Export", "No records match the selected status and search filter.");
        return;
    }
    const QString selection = statusFilter ? statusFilter->currentText() : QString("📋 All Results");
    QString safeSelection = selection;
    safeSelection.remove(QRegularExpression("[^A-Za-z0-9 -]"));
    safeSelection = safeSelection.trimmed().toLower().replace(' ', '-');
    const QString defaultName = QString("comparison-%1.pdf").arg(safeSelection);
    const QString path = QFileDialog::getSaveFileName(this, "Export visible results to PDF", defaultName, "PDF (*.pdf)");
    if (path.isEmpty()) return;
    try {
        ExcelEngine::exportPdf(visible, path);
        addAudit("PDF results exported", QString("%1: %2 records").arg(selection).arg(visible.changes.size()));
    } catch (const std::exception& error) {
        QMessageBox::critical(this, "Export", error.what());
    }
}

void MainWindow::exportAudit()
{
    if (result.changes.isEmpty()) {
        QMessageBox::information(this, "Export", "Run a comparison first.");
        return;
    }
    const Result visible = filteredResult();
    if (visible.changes.isEmpty()) {
        QMessageBox::information(this, "Export", "No records match the selected status and search filter.");
        return;
    }
    const QString selection = statusFilter ? statusFilter->currentText() : QString("📋 All Results");
    QString safeSelection = selection;
    safeSelection.remove(QRegularExpression("[^A-Za-z0-9 -]"));
    safeSelection = safeSelection.trimmed().toLower().replace(' ', '-');
    const QString defaultName = QString("comparison-%1.csv").arg(safeSelection);
    const QString path = QFileDialog::getSaveFileName(this, "Export visible results to CSV", defaultName, "CSV (*.csv)");
    if (path.isEmpty()) return;
    try {
        ExcelEngine::exportCsv(visible, path);
        addAudit("CSV results exported", QString("%1: %2 records").arg(selection).arg(visible.changes.size()));
    } catch (const std::exception& error) {
        QMessageBox::critical(this, "Export", error.what());
    }
}
void MainWindow::applyStyle(){setStyleSheet(R"(
QWidget{font:10pt "Segoe UI";background:#f8fafc;color:#0f172a}QWidget#sidebar{background:#111827}QLabel#brand{color:white;background:transparent;padding:8px}QLabel#sideReady{color:#22c55e;background:transparent}QListWidget#navigation{background:transparent;border:0;color:#e5e7eb;outline:0}QListWidget#navigation::item{padding:14px 12px;border-radius:6px;margin:2px 0}QListWidget#navigation::item:selected{background:#2563eb;color:white}QListWidget#navigation::item:hover{background:#1f2937}QLabel#pageTitle{font-size:22pt;font-weight:700;color:#172554}QGroupBox{background:white;border:1px solid #dbe3ee;border-radius:8px;margin-top:12px;padding:12px}QGroupBox::title{subcontrol-origin:margin;left:12px;padding:0 5px;font-weight:600}QLineEdit,QComboBox,QTableWidget,QTextEdit,QTextBrowser{background:white;border:1px solid #cbd5e1;border-radius:5px;padding:6px}QPushButton{background:#e2e8f0;border:0;border-radius:5px;padding:9px 14px;font-weight:600}QPushButton#modeButton{background:#e2e8f0;color:#0f172a;border:1px solid #cbd5e1}QPushButton#modeButton:checked{background:#2563eb;color:white;border:1px solid #1d4ed8}QPushButton#primaryButton{background:#2563eb;color:white}QPushButton#previewButton{background:#7c3aed;color:white}QPushButton#mergeButton{background:#059669;color:white}QPushButton:disabled{background:#cbd5e1;color:white}QLabel#hintLabel{color:#64748b}QLabel#number{font-size:20pt;font-weight:700}QGroupBox#addedCard{border-left:5px solid #22c55e}QGroupBox#deletedCard{border-left:5px solid #ef4444}QGroupBox#modifiedCard{border-left:5px solid #f59e0b}QGroupBox#unchangedCard{border-left:5px solid #64748b}QGroupBox#totalCard{border-left:5px solid #2563eb}QHeaderView::section{background:#e2e8f0;padding:8px;border:0;border-right:1px solid #cbd5e1})");}

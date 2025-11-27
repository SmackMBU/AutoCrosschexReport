CC      = gcc
CFLAGS  = -Wall -O2 -Iinclude

LFLAGS_LOCAL = -Llib

LFLAGS_SYSTEM = -lodbc32 -lxlsxwriter -lz

SRC_DIR   = src
BUILD_DIR = build

AUTO_SRC     = $(SRC_DIR)/AutoDownloadRecords.c
REPORTS_SRC  = $(SRC_DIR)/AutoReports.c

AUTO_BIN     = $(BUILD_DIR)/AutoCrosschex.exe
REPORTS_BIN  = $(BUILD_DIR)/autoreports.exe

LIBS_AUTO    = $(LFLAGS_LOCAL) -ltc-b_new_sdk -lodbc32
LIBS_REPORTS = $(LFLAGS_SYSTEM)

all: $(BUILD_DIR) $(AUTO_BIN) $(REPORTS_BIN)


$(AUTO_BIN): $(AUTO_SRC)
	$(CC) $(CFLAGS) $< -o $@ $(LIBS_AUTO)

$(REPORTS_BIN): $(REPORTS_SRC)
	$(CC) $(CFLAGS) $< -o $@ $(LIBS_REPORTS)


$(BUILD_DIR):
	if not exist "$(BUILD_DIR)" mkdir "$(BUILD_DIR)"


clean:
	if exist "$(BUILD_DIR)" rmdir /S /Q "$(BUILD_DIR)"


run_auto: $(AUTO_BIN)
	$(AUTO_BIN)

run_reports: $(REPORTS_BIN)
	$(REPORTS_BIN)

.PHONY: install

install:
	@echo "Installing Python dependencies..."
	pip install -r requirements.txt
	pip install pyinstaller
	sudo apt install python3-tk -y
	@echo ""
	@echo "✅ Installation complete!"
	@echo ""
build_exe:
	#pyinstaller --onefile --windowed --name=pdf2xlsx gui.py
	pyinstaller --onefile --windowed --name "PDF-to-Excel" main.py
	@echo ""
	@echo "✅ Build complete! Executable created in the 'dist' folder."
	@echo ""
	
clean:
	rm -rf xlsx_outputs/*.xlsx
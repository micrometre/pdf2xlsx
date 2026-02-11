.PHONY: install

install:
	@echo "Installing Python dependencies..."
	pip install -r requirements.txt
	sudo apt install python3-tk -y
	@echo ""
	@echo "✅ Installation complete!"
	@echo ""

clean:
	rm -rf xlsx_outputs
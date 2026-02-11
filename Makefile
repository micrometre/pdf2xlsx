.PHONY: install

install:
	@echo "Installing Python dependencies..."
	pip install -r requirements.txt
	@echo ""
	@echo "✅ Installation complete!"
	@echo ""

clean:
	rm -rf xlsx_outputs
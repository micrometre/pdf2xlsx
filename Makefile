.PHONY: help 

help:
	@echo ""
	@echo "Available targets:"
install:
	@echo "Installing Python dependencies..."
	pip install -r requirements.txt
	@echo "Installing Ansible collections..."
	@echo ""
	@echo "✅ Installation complete!"
	@echo ""
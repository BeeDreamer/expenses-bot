"""
Main entry point — runs both Telegram bot and Flask API concurrently
"""
import threading
import os

def run_api():
    from api import app
    port = int(os.getenv('PORT', 8000))
    print(f"🌐 API server starting on port {port}")
    app.run(host='0.0.0.0', port=port, use_reloader=False)

def run_bot():
    import bot
    bot.main()

if __name__ == '__main__':
    # Start API in background thread
    api_thread = threading.Thread(target=run_api, daemon=True)
    api_thread.start()
    
    # Run bot in main thread
    run_bot()

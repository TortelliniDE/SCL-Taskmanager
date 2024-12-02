log_messages = []  # Zentrale Liste für Log-Nachrichten

def log_and_print(message):
    """Loggt eine Nachricht und gibt sie auf der Konsole aus."""
    print(message)  # Ausgabe auf der Konsole
    log_messages.append(message)  # Speichern der Nachricht in der Log-Liste

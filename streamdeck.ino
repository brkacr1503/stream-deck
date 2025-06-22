const unsigned long KEEPALIVE_INTERVAL = 1000; // 1 second
unsigned long lastKeepAlive = 0;
const unsigned long DEBOUNCE_DELAY = 50;  // Reduced to 50ms for faster response

// Button state tracking
struct ButtonState {
  bool lastState;
  bool currentState;
  unsigned long lastDebounceTime;
  bool canSend;
};

ButtonState buttons[8];

void setup() {
  Serial.begin(9600);
  
  // Configure input pins with internal pull-up resistors
  for (int pin = 2; pin <= 9; pin++) {
    pinMode(pin, INPUT_PULLUP);
  }
  
  // Initialize button states
  for (int i = 0; i < 8; i++) {
    buttons[i].lastState = HIGH;
    buttons[i].currentState = HIGH;
    buttons[i].lastDebounceTime = 0;
    buttons[i].canSend = true;
  }
  
  // Send initial identification
  Serial.println("DECK");
}

void loop() {
  unsigned long currentMillis = millis();

  // Check for incoming messages without blocking
  while (Serial.available() > 0) {
    String input = Serial.readStringUntil('\n');
    input.trim();
    
    if (input == "TEST") {
      Serial.println("DECK");
    }
    else if (input == "PING") {
      Serial.println("PONG");
    }
  }

  // Check button states - non-blocking
  char buttonChars[] = {'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'};
  for (int i = 0; i < 8; i++) {
    // Read current button state
    bool reading = digitalRead(i + 2);

    // If the state changed, reset the debounce timer
    if (reading != buttons[i].lastState) {
      buttons[i].lastDebounceTime = currentMillis;
    }

    // Check if enough time has passed since last state change
    if ((currentMillis - buttons[i].lastDebounceTime) > DEBOUNCE_DELAY) {
      // If the state has changed:
      if (reading != buttons[i].currentState) {
        buttons[i].currentState = reading;

        // Only send on button press (LOW) and if we haven't sent recently
        if (buttons[i].currentState == LOW && buttons[i].canSend) {
          Serial.println(buttonChars[i]);
          buttons[i].canSend = false;  // Prevent multiple sends
        }
        // Reset send flag when button is released
        else if (buttons[i].currentState == HIGH) {
          buttons[i].canSend = true;
        }
      }
    }

    buttons[i].lastState = reading;
  }

  // Send periodic keepalive - non-blocking
  if (currentMillis - lastKeepAlive >= KEEPALIVE_INTERVAL) {
    Serial.println("DECK");
    lastKeepAlive = currentMillis;
  }
}
/*
  Power-Optimized LED Controller for ESP32

  This version implements multiple power-saving strategies:
  - Light sleep mode when idle
  - Reduced BLE advertising intervals
  - CPU frequency scaling
  - Optimized connection parameters
  - Automatic deep sleep during long idle periods

  Compatible with Arduino MKR WiFi 1010, Arduino Uno WiFi Rev2 board, Arduino Nano 33 IoT,
  Arduino Nano 33 BLE, or Arduino Nano 33 BLE Sense board.
*/

#include <ArduinoBLE.h>
#include <esp_pm.h>
#include <esp_wifi.h>
#include <esp_bt.h>

BLEService ledService("19B10000-E8F2-537E-4F6C-D104768A1214");
BLEByteCharacteristic switchCharacteristic("19B10001-E8F2-537E-4F6C-D104768A1214", BLERead | BLEWrite);

const int ledPin = 23;

// Power management variables
unsigned long lastActivityTime = 0;
unsigned long lastConnectionTime = 0;
const unsigned long IDLE_TIMEOUT = 300000; // 5 minutes before deep sleep
const unsigned long CONNECTION_TIMEOUT = 60000; // 1 minute without connection
const uint32_t sleepTimeMs = 100; // light sleep for 100 ms
const uint32_t deepSleepTimeSec = 30; // deep sleep for 30 seconds
bool deviceConnected = false;
bool ledState = false;

// Power management configuration
void configurePowerManagement() {
  // Enable automatic light sleep
  esp_pm_config_esp32_t pm_config;
  pm_config.max_freq_mhz = 80; // Reduce max CPU frequency from 240MHz to 80MHz
  pm_config.min_freq_mhz = 10; // Minimum frequency during light sleep
  pm_config.light_sleep_enable = true;
  esp_pm_configure(&pm_config);
  
  // Disable WiFi to save power (we only need BLE)
  esp_wifi_stop();
  esp_wifi_deinit();
  
  Serial.println("Power management configured");
}

// Optimized BLE setup
void setupOptimizedBLE() {
  // Configure BLE for lower power consumption
  BLE.setConnectionInterval(400, 800); // Slower connection interval (500-1000ms)
  BLE.setAdvertisingInterval(1600); // Slower advertising (1000ms intervals)
  
  // Set lower TX power for shorter range but better battery life
  // esp_ble_tx_power_set(ESP_BLE_PWR_TYPE_ADV, ESP_PWR_LVL_N12); // -12dBm
  
  Serial.println("BLE optimized for power saving");
}

void setup() {
  Serial.begin(115200);
  
  // Brief delay to allow serial monitor to connect
  delay(1000);
  Serial.println("Starting Power-Optimized BLE LED Controller");
  
  // Configure power management early
  configurePowerManagement();
  
  // Set LED pin to output mode
  pinMode(ledPin, OUTPUT);
  digitalWrite(ledPin, LOW);
  
  // Begin BLE initialization
  if (!BLE.begin()) {
    Serial.println("Starting Bluetooth® Low Energy module failed!");
    // Enter deep sleep if BLE fails to start
    enterDeepSleep();
  }
  
  // Set up optimized BLE parameters
  setupOptimizedBLE();
  
  // Set advertised local name and service UUID
  BLE.setLocalName("LED");
  BLE.setAdvertisedService(ledService);
  
  // Add the characteristic to the service
  ledService.addCharacteristic(switchCharacteristic);
  
  // Add service
  BLE.addService(ledService);
  
  // Set the initial value for the characteristic
  switchCharacteristic.writeValue(0);
  
  // Start advertising
  BLE.advertise();
  
  lastActivityTime = millis();
  lastConnectionTime = millis();
  
  Serial.println("BLE LED Peripheral Ready - Power Optimized");
  Serial.println("Will enter deep sleep after 5 minutes of inactivity");
}

void enterDeepSleep() {
  Serial.println("Entering deep sleep mode...");
  Serial.flush(); // Ensure message is sent
  
  // Turn off LED before sleep
  digitalWrite(ledPin, LOW);
  
  // Stop BLE advertising
  BLE.stopAdvertise();
  BLE.end();
  
  // Configure wake-up timer (wake up every 30 seconds to check for activity)
  esp_sleep_enable_timer_wakeup(deepSleepTimeSec * 1000000); // 30 seconds in microseconds
  
  // Enter deep sleep
  esp_deep_sleep_start();
}

void enterLightSleep() {
  // Light sleep maintains RAM and can wake up faster
  esp_sleep_enable_timer_wakeup(sleepTimeMs * 1000);
  esp_light_sleep_start();
}

void handleLEDControl() {
  if (switchCharacteristic.written()) {
    lastActivityTime = millis(); // Reset activity timer
    
    bool newState = switchCharacteristic.value() != 0;
    
    if (newState != ledState) {
      ledState = newState;
      
      if (ledState) {
        Serial.println("LED on");
        digitalWrite(ledPin, HIGH);
      } else {
        Serial.println("LED off");
        digitalWrite(ledPin, LOW);
      }
    }
  }
}

void checkPowerManagement() {
  unsigned long currentTime = millis();
  
  // Check for deep sleep condition (long inactivity)
  if (currentTime - lastActivityTime > IDLE_TIMEOUT) {
    Serial.println("Long idle period detected");
    enterDeepSleep();
  }
  
  // If not connected and no recent activity, enter light sleep briefly
  if (!deviceConnected && (currentTime - lastConnectionTime > 10000)) {
    // Light sleep for 100ms when no device connected and idle
    enterLightSleep();
    lastConnectionTime = millis(); // Reset to prevent immediate re-sleep
  }
}

void loop() {
  // Listen for Bluetooth® Low Energy peripherals to connect
  BLEDevice central = BLE.central();
  
  // If a central is connected to peripheral
  if (central) {
    if (!deviceConnected) {
      Serial.print("Connected to central: ");
      Serial.println(central.address());
      deviceConnected = true;
      lastActivityTime = millis();
      lastConnectionTime = millis();
    }
    
    // While the central is still connected to peripheral
    while (central.connected()) {
      handleLEDControl();
      
      // Brief delay to prevent busy-waiting and allow power management
      delay(50); // Reduced from typical 100ms for better responsiveness
      
      // Update connection time
      lastConnectionTime = millis();
    }
    
    // When the central disconnects
    Serial.print("Disconnected from central: ");
    Serial.println(central.address());
    deviceConnected = false;
    lastConnectionTime = millis();
    
    // Turn off LED when disconnected to save power
    digitalWrite(ledPin, LOW);
    ledState = false;
    switchCharacteristic.writeValue(0);
  }
  
  // Check if we should enter power saving mode
  checkPowerManagement();
  
  // Small delay to prevent busy loop when not connected
  if (!deviceConnected) {
    delay(500); // Longer delay when not connected
  }
}
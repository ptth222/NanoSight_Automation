char recieved_char;
bool autosampler_output_latch = false;

void setup() {
  pinMode(5, OUTPUT);
  pinMode(7, INPUT_PULLUP);
  Serial.begin(115200);

  digitalWrite(5, HIGH);

}

void loop() {

  if (digitalRead(7) == LOW){
    autosampler_output_latch = true;
  }


  //Look for a T sent through the serial port.
  //If found then toggle the output high to signal the autosampler.
  if (Serial.available() >0){
    recieved_char = Serial.read();
    if (recieved_char == 'T'){
      digitalWrite(5, LOW);
      delay(2000);
      digitalWrite(5, HIGH);
      Serial.println("Signal Sent To Autosampler");
    }
    else if (recieved_char == 'R'){
      if (digitalRead(7) == LOW){
        Serial.println("Signal From Autosampler Detected");
      }
      else{
        Serial.println("No Signal From Autosampler");
      } 
    }
    else if (recieved_char == 'S'){
      if (autosampler_output_latch){
        Serial.println("Autosampler Signal Latched As True");
      }
      else {
        Serial.println("Autosampler Signal Latched As False");
      }
    }
    else if (recieved_char == 'C'){
      autosampler_output_latch = false;
      Serial.println("Autosampler Signal Latch Cleared");
    }
  }
  
}

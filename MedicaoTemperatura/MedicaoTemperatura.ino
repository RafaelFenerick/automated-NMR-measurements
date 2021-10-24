//--- Declaração de variáveis
double alimentacao =5.0;            // tensão de alimentação do circuito
double resistencia_divisor =10000;  // resistência do resistor em Ohms
float tensao, resistencia;          // variáveis para cálculo

//--- Inicialização
void setup() 
{
 Serial.begin(9600);  // ativar comunicação serial
 pinMode(A0,INPUT);   // definir pino A0 como 'input'
}

//--- Leitura e transmissão de dados
void loop()
{
  // Determinação do valor de tensão com base na tensão de alimentação
  // e da saída do conversor analógico digital
  tensao = analogRead(A0)*alimentacao/1023.0;

  // Cálculo da resistência do termistor com base na tensão de alimentação,
  // a tensão no termistor e a resistência do resistor série
  resistencia = tensao/((alimentacao-tensao)/resistencia_divisor);

  Serial.print(String(resistencia) + "\n"); // envio de dados via comunicação serial
  delay(1000);                              // espera de 1s para próxima medição
}

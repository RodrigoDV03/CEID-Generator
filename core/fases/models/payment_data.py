from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from num2words import num2words


@dataclass
class PaymentData:

    # Montos base
    categoria_monto: float = 0.0
    total_pago: float = 0.0
    servicio_actualizacion: float = 0.0
    bono: float = 0.0
    disenio_examenes: float = 0.0
    examen_clasificacion: float = 0.0
    
    # Datos de control de pagos (para contratos)
    monto_total_contrato: float = 0.0
    primera_armada: float = 0.0
    segunda_armada: float = 0.0
    tercera_armada: float = 0.0
    
    def __post_init__(self):
        self.categoria_monto = float(self.categoria_monto) if self.categoria_monto else 1.0
        self.total_pago = float(self.total_pago) if self.total_pago else 0.0
        self.servicio_actualizacion = float(self.servicio_actualizacion) if self.servicio_actualizacion else 0.0
        self.bono = float(self.bono) if self.bono else 0.0
        self.disenio_examenes = float(self.disenio_examenes) if self.disenio_examenes else 0.0
        self.examen_clasificacion = float(self.examen_clasificacion) if self.examen_clasificacion else 0.0
    
    @property
    def tiene_bono(self) -> bool:
        return self.bono > 0
    
    @property
    def tiene_servicio_actualizacion(self) -> bool:
        return self.servicio_actualizacion > 0
    
    @property
    def tiene_disenio_examenes(self) -> bool:
        return self.disenio_examenes > 0
    
    @property
    def tiene_examen_clasificacion(self) -> bool:
        return self.examen_clasificacion > 0
    
    def calcular_horas_disenio(self) -> int:
        if self.categoria_monto == 0:
            return 0
        return int(round(self.disenio_examenes / self.categoria_monto))
    
    def calcular_horas_clasificacion(self) -> int:
        if self.categoria_monto == 0:
            return 0
        return int(round(self.examen_clasificacion / self.categoria_monto))
    
    def calcular_monto_total(self, es_administrativo: bool = False) -> float:
        if es_administrativo:
            return self.total_pago + self.bono
        return self.total_pago
    
    def calcular_monto_sin_actualizacion(self, es_administrativo: bool = False) -> float:
        monto_total = self.calcular_monto_total(es_administrativo)
        if es_administrativo:
            return monto_total - self.servicio_actualizacion
        return monto_total - self.servicio_actualizacion - self.bono
    
    def calcular_saldo_primera(self) -> float:
        return self.monto_total_contrato - self.primera_armada
    
    def calcular_saldo_segunda(self) -> float:
        return self.calcular_saldo_primera() - self.segunda_armada
    
    def calcular_saldo_restante(self) -> float:
        return self.monto_total_contrato - self.primera_armada - self.segunda_armada - self.tercera_armada
    
    @staticmethod
    def monto_a_letras(monto: float) -> str:
        try:
            monto_decimal = Decimal(str(monto)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            entero = int(monto_decimal)
            centavos = int((monto_decimal - Decimal(entero)) * 100)
            return f"{num2words(entero, lang='es')} y {centavos:02d}/100 soles"
        except Exception:
            return "N/A"
    
    def formatear_monto(self, monto: float) -> str:
        return f"S/. {monto:,.2f}"
    
    def formatear_con_letras(self, monto: float) -> str:
        return f"{self.formatear_monto(monto)} ({self.monto_a_letras(monto)})"

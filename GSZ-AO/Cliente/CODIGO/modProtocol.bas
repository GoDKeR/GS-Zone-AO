Attribute VB_Name = "modProtocol"
'**************************************************************
' modProtocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @file     modProtocol.bas
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20060517

Option Explicit

''
' TODO : /BANIP y /UNBANIP ya no trabajan con nicks. Esto lo puede mentir en forma local el cliente con un paquete a NickToIp

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Private Type tFont
    Red As Byte
    Green As Byte
    Blue As Byte
    bold As Boolean
    italic As Boolean
End Type

Public Enum eMensajes 'By TwIsT
    Mensaje001 ' "Comercio cancelado por el otro usuario." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje002 ' "Has terminado de descansar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje003 ' "¡¡¡Estás obstruyendo la vía pública, muévete o serás encarcelado!!!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje004 ' "Has sido expulsado del clan. ¡El clan ha sumado un punto de antifacción!" *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje005 ' "¡¡Estás muerto!!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje006 ' "Estás demasiado lejos del vendedor." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje007 ' "El sacerdote no puede curarte debido a que estás demasiado lejos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje008 ' "Estas demasiado lejos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje009 ' "La puerta esta cerrada con llave." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje010 ' "La puerta está cerrada con llave." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje011 ' "Estás demasiado lejos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje012 ' "No puedes hacer fogatas en zona segura." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje013 ' "Has prendido la fogata." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje014 ' "La ley impide realizar fogatas en las ciudades." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje015 ' "No has podido hacer fuego." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje016 ' "¡Has sido liberado!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje017 ' "El usuario no está online." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje018 ' "No puedes banear a al alguien de mayor jerarquía." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje019 ' "El personaje ya se encuentra baneado." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje020 ' "La mascota no atacará a ciudadanos si eres miembro del ejército real o tienes el seguro activado." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje021 ' "No tienes suficiente dinero." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje022 ' "No puedes cargar mas objetos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje023 ' "Lo siento, no estoy interesado en este tipo de objetos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje024 ' "Las armaduras del ejército real sólo pueden ser vendidas a los sastres reales." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje025 ' "Las armaduras de la legión oscura sólo pueden ser vendidas a los sastres del demonio." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje026 ' "No puedes vender ítems." *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje027 ' "Mapa exclusivo para newbies." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje028 ' "Mapa exclusivo para miembros del ejército real." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje029 ' "Mapa exclusivo para miembros de la legión oscura." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje030 ' "Solo se permite entrar al mapa si eres miembro de alguna facción." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje031 ' "Comercio cancelado. El otro usuario se ha desconectado." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje032 ' "¡¡Estás muriendo de frío, abrigate o morirás!!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje033 ' "¡¡Has muerto de frío!!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje034 ' "¡¡Quitate de la lava, te estás quemando!!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje035 ' "¡¡Has muerto quemado!!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje036 ' "Recuperas tu apariencia normal." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje037 ' "Has vuelto a ser visible." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje038 ' "Estás envenenado, si no te curas morirás." *-*  FontTypeNames.FONTTYPE_VENENO
    Mensaje039 ' "Has sanado." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje040 ' "Gracias por jugar Argentum Online" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje041 ' "No puedes tirar objetos newbie." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje042 ' "No hay espacio en el piso." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje043 ' "¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU BARCA!" *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje044 ' "No puedes cargar más objetos." *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje045 ' "No hay nada aquí." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje046 ' "Sólo los newbies pueden usar este objeto." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje047 ' "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje048 ' "Sólo los newbies pueden usar estos objetos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje049 ' "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje050 ' "Antes de usar la herramienta deberías equipartela." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje051 ' "Debes tener equipada la herramienta para trabajar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje052 ' "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo. " *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje053 ' "¡¡Debes esperar unos momentos para tomar otra poción!!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje054 ' "Te has curado del envenenamiento." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje055 ' "Sientes un gran mareo y pierdes el conocimiento." *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje056 ' "Has abierto la puerta." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje057 ' "La llave no sirve." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje058 ' "Has cerrado con llave la puerta." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje059 ' "No está cerrada." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje060 ' "No hay agua allí." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje061 ' "Estás demasiado hambriento y sediento." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje062 ' "No tienes conocimientos de las Artes Arcanas." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje063 ' "No hay peligro aquí. Es zona segura." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje064 ' "Sólo miembros del ejército real pueden usar este cuerno." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje065 ' "Sólo miembros de la legión oscura pueden usar este cuerno." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje066 ' "Para recorrer los mares debes ser nivel 25 o superior." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje067 ' "Para recorrer los mares debes ser nivel 20 o superior." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje068 ' "¡Debes aproximarte al agua para usar el barco!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje069 ' "Tu carisma y liderazgo no son suficientes para liderar una party." *-*  FontTypeNames.FONTTYPE_PARTY
    Mensaje070 ' "Por el momento no se pueden crear más parties." *-*  FontTypeNames.FONTTYPE_PARTY
    Mensaje071 ' "La party está llena, no puedes entrar." *-*  FontTypeNames.FONTTYPE_PARTY
    Mensaje072 ' "¡Has formado una party!" *-*  FontTypeNames.FONTTYPE_PARTY
    Mensaje073 ' "No puedes hacerte líder." *-*  FontTypeNames.FONTTYPE_PARTY
    Mensaje074 ' "¡Te has convertido en líder de la party!" *-*  FontTypeNames.FONTTYPE_PARTY
    Mensaje075 ' "No tienes suficientes puntos de liderazgo para liderar una party." *-*  FontTypeNames.FONTTYPE_PARTY
    Mensaje076 ' "Ya perteneces a una party." *-*  FontTypeNames.FONTTYPE_PARTY
    Mensaje077 ' "Ya perteneces a una party, escribe /SALIRPARTY para abandonarla" *-*  FontTypeNames.FONTTYPE_PARTY
    Mensaje078 ' "El fundador decidirá si te acepta en la party." *-*  FontTypeNames.FONTTYPE_PARTY
    Mensaje079 ' "Para ingresar a una party debes hacer click sobre el fundador y luego escribir /PARTY" *-*  FontTypeNames.FONTTYPE_PARTY
    Mensaje080 ' "No eres miembro de ninguna party." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje081 ' "¡No eres el líder de tu party!" *-*  FontTypeNames.FONTTYPE_PARTY
    Mensaje082 ' "¡Está muerto, no puedes aceptar miembros en ese estado!" *-*                      FontTypeNames.FONTTYPE_PARTY
    Mensaje083 ' "¡No se ha hecho el cambio de mando!" *-*  FontTypeNames.FONTTYPE_PARTY
    Mensaje084 ' "¡Está muerto!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje085 ' "No podés tener mas objetos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje086 ' "No tienes mas espacio en el banco!!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje087 ' "El banco no puede cargar tantos objetos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje088 ' "El centinela intenta llamar tu atención. ¡Respóndele rápido!" *-*  FontTypeNames.FONTTYPE_CENTINELA
    Mensaje089 ' "El centinela intenta llamar tu atención. ¡Respondele rápido!" *-*  FontTypeNames.FONTTYPE_CENTINELA
    Mensaje090 ' "¡¡¡Has sido expulsado del ejército real!!!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje091 ' "¡¡¡Te has retirado del ejército real!!!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje092 ' "¡¡¡Has sido expulsado de la Legión Oscura!!!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje093 ' "¡¡¡Te has retirado de la Legión Oscura!!!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje094 ' "Hoy es la votación para elegir un nuevo líder para el clan." *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje095 ' "La elección durará 24 horas, se puede votar a cualquier miembro del clan." *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje096 ' "Para votar escribe /VOTO NICKNAME." *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje097 ' "Sólo se computará un voto por miembro. Tu voto no puede ser cambiado." *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje098 ' "Error, el clan no existe." *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje099 ' "No perteneces a ningún clan." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje100 ' "No eres el líder de tu clan." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje101 ' "El personaje no es ni aspirante ni miembro del clan." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje102 ' "No está permitido arrojar objetos al suelo en zonas seguras.", FontTypeNames.FONTTYPE_CITIZEN
    Mensaje103 ' "No tienes espacio para más hechizos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje104 ' "Ya sabes el hechizo." *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje105 ' "No puedes lanzar hechizos estando muerto." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje106 ' "No posees un báculo lo suficientemente poderoso para poder lanzar el conjuro." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje107 ' "No puedes lanzar este conjuro sin la ayuda de un báculo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje108 ' "No tienes suficientes puntos de magia para lanzar este hechizo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje109 ' "Estás muy cansado para lanzar este hechizo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje110 ' "Estás muy cansada para lanzar este hechizo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje111 ' "Debes poseer toda tu maná para poder lanzar este hechizo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje112 ' "Debes poseer alguna mascota para poder lanzar este hechizo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje113 ' "No tienes suficiente maná." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje114 ' "No puedes invocar criaturas en zona segura." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje115 ' "No puedes lanzar hechizos si estás en consulta." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje116 ' "Estás demasiado lejos para lanzar este hechizo." *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje117 ' "Este hechizo actúa sólo sobre usuarios." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje118 ' "Este hechizo sólo afecta a los npcs." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje119 ' "Target inválido." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje120 ' "¡El usuario está muerto!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje121 ' "¡El hechizo no tiene efecto!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje122 ' "¡No puedes hacerte invisible mientras te encuentras saliendo!" *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje123 ' "¡La invisibilidad no funciona aquí!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje124 ' "Ya te encuentras mimetizado. El hechizo no ha tenido efecto." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje125 ' "No puedes atacarte a vos mismo." *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje126 ' " ¡El hechizo no tiene efecto!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje127 ' "¡El espíritu no tiene intenciones de regresar al mundo de los vivos!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje128 ' "¡Revivir no está permitido aquí! Retirate de la Zona si deseas utilizar el Hechizo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje129 ' "No puedes resucitar si no tienes tu barra de energía llena." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje130 ' "Necesitas un báculo mejor para lanzar este hechizo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje131 ' "Necesitas un instrumento mágico para devolver la vida." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje132 ' "¡Los Dioses te sonríen, has ganado 500 puntos de nobleza!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje133 ' "El esfuerzo de resucitar fue demasiado grande." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje134 ' "El esfuerzo de resucitar te ha debilitado." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje135 ' "Tu viaje ha sido cancelado." *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje136 ' "El NPC es inmune a este hechizo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje137 ' "Sólo puedes remover la parálisis de los Guardias si perteneces a su facción." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje138 ' "Solo puedes remover la parálisis de los NPCs que te consideren su amo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje139 ' "Solo puedes remover la parálisis de los Guardias si perteneces a su facción." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje140 ' "Este NPC no está paralizado" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje141 ' "El NPC es inmune al hechizo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje142 ' "Sólo los druidas pueden mimetizarse con criaturas." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje143 ' "No puedes lanzar este hechizo a un muerto." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje144 ' "No puedes ayudar usuarios mientras estas en consulta." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje145 ' "Los miembros del ejército real no pueden ayudar a los criminales." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje146 ' "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje147 ' "Los miembros de la legión oscura no pueden ayudar a los ciudadanos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje148 ' "Los miembros del ejército real no pueden ayudar a ciudadanos en estado atacable." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje149 ' "Para ayudar ciudadanos en estado atacable debes sacarte el seguro, pero te puedes volver criminal." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje150 ' "No puedes mover el hechizo en esa dirección." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje151 ' "¡Has matado a la criatura!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje152 ' "¡¡La criatura te ha envenenado!!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje153 ' "¡Has subido de nivel!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje154 ' "¡Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearás! Por esta razón, hasta tanto no te enlistes en la facción bajo la cual tu clan está alineado, estarás excluído del mismo." *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje155 ' "Debes abandonar el Dungeon Newbie." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje156 ' "(CUERPO) Mín Def/Máx Def: 0" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje157 ' "(CABEZA) Mín Def/Máx Def: 0" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje158 ' "Status: Líder" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje159 ' "Fue ejército real" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje160 ' "Fue legión oscura" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje161 ' "Para poder entrenar un skill debes asignar los 10 skills iniciales." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje162 ' "¡Has ganado 50 puntos de experiencia!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje163 ' "Tus mascotas no pueden transitar este mapa." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje164 ' "Pierdes el control de tus mascotas invocadas." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje165 ' "No se permiten mascotas en zona segura. Éstas te esperarán afuera." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje166 ' "Tu mascota no pueden transitar este sector del mapa, intenta invocarla en otra parte." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje167 ' "¡Has recuperado tu apariencia normal!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje168 ' "/salir cancelado." *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje169 ' "Personaje Inexistente" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje170 ' "Debes estar muerto para poder utilizar este comando." *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje171 ' "No puedes robar npcs en zonas seguras." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje172 ' "No puedes atacar otra criatura con dueño hasta que haya terminado tu castigo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje173 ' "El rey pretoriano te ha vuelto ciego " *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje174 ' "A la distancia escuchas las siguientes palabras: ¡Cobarde, no eres digno de luchar conmigo si escapas! " *-*  FontTypeNames.FONTTYPE_VENENO
    Mensaje175 ' "El rey pretoriano te ha vuelto estúpido." *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje176 ' "¡Has sido detectado!" *-*  FontTypeNames.FONTTYPE_VENENO
    Mensaje177 ' "Comienzas a hacerte visible." *-*  FontTypeNames.FONTTYPE_VENENO
    Mensaje178 ' "Ya te encuentras en tu hogar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje179 ' "No puedes usar este comando aquí." *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje180 ' "Debes estar muerto para utilizar este comando." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje181 ' "¡Has vuelto a ser visible!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje182 ' "¡¡Estás muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje183 ' "Usuario inexistente." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje184 ' "No puedes susurrarle a los Dioses y Admins." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje185 ' "No puedes susurrarle a los GMs." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje186 ' "Estás muy lejos del usuario." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje187 ' "Dejas de meditar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje188 ' "Has dejado de descansar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje189 ' "No puedes moverte porque estás paralizado." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje190 ' "No puedes usar así este arma." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje191 ' "No puedes tomar ningún objeto." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje192 ' "Has dejado de comerciar." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje193 ' "Has rechazado la oferta del otro usuario." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje194 ' "No puedes ocultarte si estás en consulta." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje195 ' "No puedes ocultarte si estás navegando." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje196 ' "Ya estás oculto." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje197 ' "No tienes municiones." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje198 ' "Estás muy cansado para luchar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje199 ' "Estás muy cansada para luchar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje200 ' "Estás demasiado lejos para atacar." *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje201 ' "¡No puedes atacarte a vos mismo!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje202 ' "Una fuerza oscura te impide canalizar tu energía." *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje203 ' "¡Primero selecciona el hechizo que quieres lanzar!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje204 ' "No puedes pescar desde donde te encuentras." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje205 ' "Estás demasiado lejos para pescar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje206 ' "No hay agua donde pescar. Busca un lago, río o mar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje207 ' "No puedes robar aquí." *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje208 ' "¡No hay a quien robarle!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje209 ' "¡No puedes robar en zonas seguras!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje210 ' "Deberías equiparte el hacha." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje211 ' "No puedes talar desde allí." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje212 ' "El hacha utilizado no es suficientemente poderosa." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje213 ' "No hay ningún árbol ahí." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje214 ' "Ahí no hay ningún yacimiento." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje215 ' "No puedes domar una criatura que está luchando con un jugador." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje216 ' "No puedes domar a esa criatura." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje217 ' "¡No hay ninguna criatura allí!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje218 ' "No tienes más minerales." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje219 ' "Ahí no hay ninguna fragua." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje220 ' "Ahí no hay ningún yunque." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje221 ' "¡Primero selecciona el hechizo!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje222 ' "No estás comerciando." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje223 ' "Propuesta de paz enviada." *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje224 ' "Propuesta de alianza enviada." *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje225 ' "El personaje no ha mandado solicitud, o no estás habilitado para verla." *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje226 ' "No puedes expulsar ese personaje del clan." *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje227 ' "No puedes salir estando paralizado." *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje228 ' "Comercio cancelado." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje229 ' "Dejas el clan." *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje230 ' "Tú no puedes salir de este clan." *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje231 ' "Primero tienes que seleccionar un NPC, haz click izquierdo sobre él." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje232 ' "¡¡Estás muerto!! Solo puedes usar ítems cuando estás vivo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje233 ' "Te acomodás junto a la fogata y comienzas a descansar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje234 ' "Te levantas." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje235 ' "No hay ninguna fogata junto a la cual descansar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje236 ' "¡¡Estás muerto!! Sólo puedes meditar cuando estás vivo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje237 ' "Sólo las clases mágicas conocen el arte de la meditación." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje238 ' "Maná restaurado." *-*  FontTypeNames.FONTTYPE_VENENO
    Mensaje239 ' "El sacerdote no puede resucitarte debido a que estás demasiado lejos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje240 ' "¡¡Has sido resucitado!!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje241 ' "Primero tienes que seleccionar un usuario, haz click izquierdo sobre él." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje242 ' "No puedes iniciar el modo consulta con otro administrador." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje243 ' "Has terminado el modo consulta." *-*  FontTypeNames.FONTTYPE_INFOBOLD
    Mensaje244 ' "Has iniciado el modo consulta." *-*  FontTypeNames.FONTTYPE_INFOBOLD
    Mensaje245 ' "¡¡Has sido curado!!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje246 ' "Ya estás comerciando." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje247 ' "¡¡No puedes comerciar con los muertos!!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje248 ' "¡¡No puedes comerciar con vos mismo!!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje249 ' "Estás demasiado lejos del usuario." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje250 ' "No puedes comerciar con el usuario en este momento." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje251 ' "Primero haz click izquierdo sobre el personaje." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje252 ' "Debes acercarte más." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje253 ' "No puedes compartir npcs con administradores!!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje254 ' "Solo puedes compartir npcs con miembros de tu misma facción!!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje255 ' "No puedes compartir npcs con criminales!!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje256 ' "No pertences a ningún clan." *-*  FontTypeNames.FONTTYPE_GUILDMSG
    Mensaje257 ' "Su solicitud ha sido enviada." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje258 ' "El mensaje ha sido entregado, ahora sólo debes esperar que se desocupe algún GM." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje259 ' "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje260 ' "No puedes cambiar la descripción estando muerto." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje261 ' "La descripción tiene caracteres inválidos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje262 ' "La descripción ha cambiado." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje263 ' "Voto contabilizado." *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje264 ' "No puedes ver las penas de los administradores." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje265 ' "Sin prontuario.." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje266 ' "Debes especificar una contraseña nueva, inténtalo de nuevo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje267 ' "La contraseña actual proporcionada no es correcta. La contraseña no ha sido cambiada, inténtalo de nuevo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje268 ' "La contraseña fue cambiada con éxito." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje269 ' "¡No perteneces a ninguna facción!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje270 ' "Denuncia enviada, espere.." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje271 ' "¡Ya has fundado un clan, no puedes fundar otro!" *-*  FontTypeNames.FONTTYPE_INFOBOLD
    Mensaje272 ' "Alineación inválida." *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje273 ' "No puedes incorporar a tu party a personajes de mayor jerarquía." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje274 ' "No hay reales conectados." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje275 ' "No hay Caos conectados." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje276 ' "Usuario offline." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje277 ' "Todos los lugares están ocupados." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje278 ' "Comentario salvado..." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje279 ' "Npcs Hostiles en mapa: " *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje280 ' "No hay NPCS Hostiles." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje281 ' "Otros Npcs en mapa: " *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje282 ' "No hay más NPCS." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje283 ' "Usuario silenciado." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje284 ' "Usuario des silenciado." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje285 ' "No perteneces a ningún grupo!" *-*  FontTypeNames.FONTTYPE_INFOBOLD
    Mensaje286 ' "No hay usuarios trabajando." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje287 ' "No hay usuarios ocultandose." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje288 ' "Utilice /carcel nick@motivo@tiempo" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje289 ' "No puedes encarcelar a administradores." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje290 ' "No puedés encarcelar por más de 60 minutos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje291 ' "Los consejeros no pueden usar este comando en el mapa pretoriano." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje292 ' "Antes debes hacer click sobre el NPC." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje293 ' "Utilice /advertencia nick@motivo" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje294 ' "No puedes advertir a administradores." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje295 ' "Estás intentando editar un usuario inexistente." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje296 ' "Clase desconocida. Intente nuevamente." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje297 ' "Skill Inexistente!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje298 ' "Genero desconocido. Intente nuevamente." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje299 ' "Raza desconocida. Intente nuevamente." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje300 ' "Comando no permitido." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje301 ' "Usuario offline, buscando en charfile." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje302 ' "Usuario offline. Leyendo charfile... " *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje303 ' "Usuario offline. Leyendo del charfile..." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje304 ' "No hay GMs Online." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje305 ' "Sólo se permite perdonar newbies." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje306 ' "No puedes echar a alguien con jerarquía mayor a la tuya." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje307 ' "¿¿Estás loco?? ¿¿Cómo vas a piñatear un gm?? :@" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje308 ' "No está online." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje309 ' "Charfile inexistente (no use +)." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje310 ' "El jugador no está online." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje311 ' "No puedes invocar a dioses y admins." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje312 ' "No hay ningún personaje con ese nick." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje313 ' "Hay un objeto en el piso en ese lugar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje314 ' "No puedes crear un teleport que apunte a la entrada de otro." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje315 ' "Haz click sobre un personaje antes." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje316 ' "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje317 ' "Usuario offline" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje318 ' "Usuario offline, echando de los consejos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje319 ' "Has sido echado del consejo de Banderbill." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje320 ' "Has sido echado del Concilio de las Sombras." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje321 ' "El personaje no está online." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje322 ' "¡¡ATENCIÓN: FUERON CREADOS ***100*** ÍTEMS, TIRE Y /DEST LOS QUE NO NECESITE!!" *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje323 ' "No puede destruir teleports así. Utilice /DT." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje324 ' "Utilice /borrarpena Nick@NumeroDePena@NuevaPena" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje325 ' "Pena modificada." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje326 ' "No hay ningún objeto en slot seleccionado." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje327 ' "Slot Inválido." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje328 ' "Npcs.dat recargado." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje329 ' "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje330 ' "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje331 ' "Igualmente, el único útil es 'NIEVE' ya que al ingresarlo, la gente muere de frío en el mapa." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje332 ' "Igualmente, el único útil es 'DUNGEON' ya que al ingresarlo, NO se sentirá el efecto de la lluvia en este mapa." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje333 ' "Mapa Guardado." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje334 ' "Usar: /ANAME origen@destino" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje335 ' "El Pj está online, debe salir para hacer el cambio." *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje336 ' "Transferencia exitosa." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje337 ' "El nick solicitado ya existe." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje338 ' "usar /AEMAIL <pj>-<nuevomail>" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje339 ' "usar /APASS <pjsinpass>@<pjconpass>" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje340 ' "Servidor habilitado para todos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje341 ' "Servidor restringido a administradores." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje342 ' "No pertenece a ningún clan o es fundador." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje343 ' "Expulsado." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje344 ' "Se ha cambiado el MOTD con éxito." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje345 ' "¡No puedes modificar esa información desde aquí!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje346 ' "No existe la llave y/o clave" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje347 ' "Debes matar al resto del ejército antes de atacar al rey!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje348 ' "No puedes atacar mascotas en zona segura." *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje349 ' "No puedes atacar a este NPC." *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje350 ' "Estás muy lejos para disparar." *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje351 ' "No puedes atacar a un espíritu." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje352 ' "No puedes atacar usuarios mientras estas en consulta." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje353 ' "No puedes atacar usuarios mientras estan en consulta." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje354 ' "El ser es demasiado poderoso." *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje355 ' "Los soldados del ejército real tienen prohibido atacar ciudadanos." *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje356 ' "Los miembros de la legión oscura tienen prohibido atacarse entre sí." *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje357 ' "No puedes atacar ciudadanos, para hacerlo debes desactivar el seguro." *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje358 ' "¡Huye de la ciudad! Estás siendo atacado y no podrás defenderte." *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje359 ' "Esta es una zona segura, aquí no puedes atacar a otros usuarios." *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje360 ' "No puedes pelear aquí." *-*  FontTypeNames.FONTTYPE_WARNING
    Mensaje361 ' "No puedes atacar npcs mientras estas en consulta." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje362 ' "No puedes atacar esta criatura." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje363 ' "No puedes atacar Guardias del Caos siendo de la legión oscura." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje364 ' "No puedes atacar Guardias Reales siendo del ejército real." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje365 ' "Para poder atacar Guardias Reales debes quitarte el seguro." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje366 ' "¡Atacaste un Guardia Real! Eres un criminal." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje367 ' "Los miembros del ejército real no pueden atacar npcs no hostiles." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje368 ' "Para atacar a este NPC debes quitarte el seguro." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje369 ' "Atacaste un NPC no-hostil. Continúa haciéndolo y te podrás convertir en criminal." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje370 ' "Los miembros del ejército real no pueden atacar mascotas de ciudadanos." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje371 ' "Para atacar mascotas de ciudadanos debes quitarte el seguro." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje372 ' "Has atacado la Mascota de un ciudadano. Eres un criminal." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje373 ' "Los miembros de la legión oscura no pueden atacar mascotas de otros legionarios. " *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje374 ' "Los miembros del Ejército Real no pueden paralizar criaturas ya paralizadas pertenecientes a otros miembros del Ejército Real" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje375 ' "Para paralizar criaturas ya paralizadas pertenecientes a ciudadanos debes quitarte el seguro." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje376 ' "Has paralizado la criatura de un ciudadano, ahora eres atacable por él." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje377 ' "Los miembros de la legión oscura no pueden paralizar criaturas ya paralizadas por otros legionarios." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje378 ' "Los miembros del Ejército Real no pueden atacar criaturas pertenecientes a otros miembros del Ejército Real" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje379 ' "Para atacar criaturas ya pertenecientes a ciudadanos debes quitarte el seguro." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje380 ' "Has atacado a la criatura de un ciudadano, ahora eres atacable por él." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje381 ' "Para atacar criaturas pertenecientes a ciudadanos debes quitarte el seguro." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje382 ' "Los miembros de la Legión Oscura no pueden atacar criaturas de otros legionarios. " *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje383 ' "Debes matar al resto del ejército antes de atacar al rey." *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje384 ' "Comercio cancelado por el otro usuario" *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje385 ' "Servidor> Por favor espera algunos segundos, el WorldSave está ejecutándose." *-*  FontTypeNames.FONTTYPE_SERVER
    Mensaje386 ' "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde." *-*  FontTypeNames.FONTTYPE_SERVER
    Mensaje387 ' "Tu estado no te permite entrar al clan." *-*  FontTypeNames.FONTTYPE_GUILD
    Mensaje388 ' "¡Te has escondido entre las sombras!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje389 ' "¡No has logrado esconderte!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje390 ' "No tienes suficientes conocimientos para usar este barco." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje391 ' "No tienes conocimientos de minería suficientes para trabajar este mineral." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje392 ' "No tienes los conocimientos suficientes en herrería para fundir este objeto." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje393 ' "No tienes suficiente madera." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje394 ' "No tienes suficiente madera élfica." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje395 ' "No tienes suficientes lingotes de hierro." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje396 ' "No tienes suficientes lingotes de plata." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje397 ' "No tienes suficientes lingotes de oro." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje398 ' "No tienes suficientes materiales." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje399 ' "No tienes suficiente energía." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje400 ' "Debes tener equipado el serrucho para trabajar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje401 ' "No tienes suficientes minerales para hacer un lingote." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje402 ' "Debes equiparte el martillo de herrero." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje403 ' "No tienes suficientes skills." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje404 ' "Has mejorado el arma!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje405 ' "Has mejorado el escudo!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje406 ' "Has mejorado el casco!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje407 ' "Has mejorado la armadura!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje408 ' "Debes equiparte el serrucho." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje409 ' "Has mejorado la flecha!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje410 ' "Has mejorado el barco!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje411 ' "Ya domaste a esa criatura." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje412 ' "La criatura ya tiene amo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje413 ' "No puedes domar más de dos criaturas del mismo tipo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje414 ' "La criatura te ha aceptado como su amo." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje415 ' "No has logrado domar la criatura." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje416 ' "No puedes controlar más criaturas." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje417 ' "Necesitas clickear sobre leña para hacer ramitas." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje418 ' "Estás demasiado lejos para prender la fogata." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje419 ' "No puedes hacer fogatas estando muerto." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje420 ' "Necesitas por lo menos tres troncos para hacer una fogata." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje421 ' "No has podido hacer la fogata." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje422 ' "¡Has pescado un lindo pez!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje423 ' "¡No has pescado nada!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje424 ' "¡Has pescado algunos peces!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje425 ' "Debes quitarte el seguro para robarle a un ciudadano." *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje426 ' "Los miembros del ejército real no tienen permitido robarle a ciudadanos." *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje427 ' "No puedes robar a otros miembros de la legión oscura." *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje428 ' "Estás muy cansado para robar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje429 ' "Estás muy cansada para robar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje430 ' "¡No has logrado robar nada!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje431 ' "No has logrado robar ningún objeto." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje432 ' "¡No has logrado apuñalar a tu enemigo!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje433 ' "¡Has conseguido algo de leña!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje434 ' "¡No has obtenido leña!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje435 ' "¡Has extraido algunos minerales!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje436 ' "¡No has conseguido nada!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje437 ' "Has terminado de meditar." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje438 ' "Has logrado desequipar el escudo de tu oponente!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje439 ' "¡Tu oponente te ha desequipado el escudo!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje440 ' "Has logrado desarmar a tu oponente!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje441 ' "¡Tu oponente te ha desarmado!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje442 ' "Has logrado desequipar el casco de tu oponente!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje443 ' "¡Tu oponente te ha desequipado el casco!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje444 ' "Tu oponente no tiene equipado items!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje445 ' "No has logrado desequipar ningún item a tu oponente!" *-*  FontTypeNames.FONTTYPE_FIGHT
    Mensaje446 ' "Tu golpe ha dejado inmóvil a tu oponente" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje447 ' "¡El golpe te ha dejado inmóvil!" *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje448 ' "Operación realizada con exito!!" *-*  FontTypeNames.FONTTYPE_INFO -- GSZ
    Mensaje449 ' "El usuario no se encuentra en el listado solicitado." *-*  FontTypeNames.FONTTYPE_INFO -- GSZ
    Mensaje450 ' "Para recorrer los mares debes ser nivel 20 y además tu skill en pesca debe ser 100." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje451 ' "Para recorrer los mares debes ser nivel 20 o superior." *-*  FontTypeNames.FONTTYPE_INFO
    Mensaje452 ' "No puedes comerciar en este momento" *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje453 ' "¡Los miembros del staff no pueden crear partys!" *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje454 ' "¡Los miembros del staff no pueden unirse a partys!" *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje455 ' "Invocar no está permitido aquí! Retirate de la Zona si deseas utilizar el Hechizo." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje456 ' "¡¡¡Tu estado no te permite permanecer en el mapa!!!" *-*  FontTypeNames.FONTTYPE_INFOBOLD
    Mensaje457 ' "Has vuelto a ser visible ya que no esta permitida la invisibilidad en este mapa." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje458 ' "Has vuelto a ser visible ya que no esta permitido ocultarse en este mapa." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje459 ' "¡Ocultarse no funciona aquí!" *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje460 ' "No hay un yacimiento de peces donde pescar." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje461 ' "No puedes pescar desde allí." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje462 ' "No puedes transportar dioses o admins." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje463 ' "Posición inválida." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje464 ' "No puedes ver está información de un dios o administrador." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje465 ' "Servidor.ini actualizado correctamente." *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje466 ' "No puedes sumonear miembros del clan pretoriano de esta forma, utiliza /CrearClanPretoriano."  *-*   FontTypeNames.FONTTYPE_WARNING
    Mensaje467 ' "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!" *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje468 ' "¡¡¡No puedes robar a usuarios en consulta!!!" *-*  FontTypeNames.FONTTYPE_TALK
    Mensaje469 ' "¡¡Comercio cancelado, te están robando!!"  *-*   FontTypeNames.FONTTYPE_WARNING
    Mensaje470 ' "¡¡¡No puedes tirar este tipo de objeto!!!", FontTypeNames.FONTTYPE_FIGHT
    Mensaje471 ' "No puedes vender este tipo de objeto.", FontTypeNames.FONTTYPE_INFO
    Mensaje472 ' "Tu anillo rechaza los efectos del hechizo inmobilizar.", FontTypeNames.FONTTYPE_FIGHT
    Mensaje473 ' "Tu anillo rechaza los efectos de la turbación.", FontTypeNames.FONTTYPE_FIGHT
    Mensaje474 ' "Tu anillo rechaza los efectos de la ceguera.", FontTypeNames.FONTTYPE_FIGHT
    Mensaje475 ' "Tu anillo rechaza los efectos de la paralisis.", FontTypeNames.FONTTYPE_FIGHT
    Mansaje476 ' "El hechizo no pertenece a tu clase.", FontTypeNames.FONTTYPE_INFO
    Mansaje477 ' "El hechizo no pertenece a tu raza.", FontTypeNames.FONTTYPE_INFO
    Mensaje478 ' "Debes hacer click sobre un personaje.",  FontTypeNames.FONTTYPE_WARNING
    Mensaje479 ' "Has conseguido algo de agua." *-*  FontTypeNames.FONTTYPE_INFO
End Enum 'By TwIsT

Private Enum ServerPacketID
    Logged                  ' LOGGED
    InfoTorneo
    ClientConfig            ' CLIENTCFG - GSZAO especial para opciones en el cliente
    CreateParticleInChar    ' CPCHAR - GSZAO crea particulas en chars
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    CreateRenderValue       ' CDMG - GSZAO
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    UserOfferConfirm
    CommerceChat
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateBankGold
    UpdateExp               ' ASE
    ChangeMap               ' CM
    PosUpdate               ' PU
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+
    ShowMessageBox          ' !!
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterChangeNick
    CharacterMove           ' MP, +, * and _ '
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayMIDI                ' TM
    PlayWave                ' TW
    guildList               ' GL
    AreaChanged             ' CA
    PauseToggle             ' BKW
    RainToggle              ' LLU
    CreateFX                ' CFX
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01
    ChangeInventorySlot     ' CSI
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    Atributes               ' ATR
    BlacksmithWeapons       ' LAH
    BlacksmithArmors        ' LAR
    CarpenterObjects        ' OBR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    Fame                    ' FAMA
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    AddForumMsg             ' FMSG
    ShowForumForm           ' MFOR
    SetInvisible            ' NOVER
    DiceRoll                ' DADOS
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    guildNews               ' GUILDNE
    OfferDetails            ' PEACEDE & ALLIEDE
    AlianceProposalsList    ' ALLIEPR
    PeaceProposalsList      ' PEACEPR
    CharacterInfo           ' CHRINFO
    GuildLeaderInfo         ' LEADERI
    GuildMemberInfo
    GuildDetails            ' CLANDET
    ShowGuildFundationForm  ' SHOWFUN
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    TradeOK                 ' TRANSOK
    BankOK                  ' BANCOOK
    ChangeUserTradeSlot     ' COMUSUINV
    Pong
    UpdateTagAndStatus
    FormYesNo               ' GSZAO
    Mensajes                ' GSZAO
    Online                  ' GSZAO
    CreateParticle
    
    'GM messages
    SpawnList               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    ' 0.13.3
    ShowDenounces
    RecordList
    RecordDetails
    
    ShowGuildAlign
    ShowPartyForm
    UpdateStrenghtAndDexterity
    UpdateStrenght
    UpdateDexterity
    AddSlots
    MultiMessage
    StopWorking
    CancelOfferItem
    
    QuestDetails            ' GSZAO
    QuestListSend           ' GSZAO
    
    UserDeath
End Enum

Private Enum ClientPacketID
    TorneoEventoInfo
    TorneoEvento
    
    LoginExistingChar       'OLOGIN
    ThrowDices              'TIRDAD
    LoginNewChar            'NLOGIN
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle
    RequestGuildLeaderInfo  'GLINFO
    RequestAtributes        'ATR
    RequestFame             'FAMA
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    UserCommerceConfirm
    CommerceChat
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem                 'USA
    CraftBlacksmith         'CNS
    CraftCarpenter          'CNC
    WorkLeftClick           'WLC
    CreateNewGuild          'CIG
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    BankDeposit             'DEPO
    ForumPost               'DEMSG
    MoveSpell               'DESPHE
    MoveBank
    ClanCodexUpdate         'DESCOD
    UserCommerceOffer       'OFRECER
    GuildAcceptPeace        'ACEPPEAT
    GuildRejectAlliance     'RECPALIA
    GuildRejectPeace        'RECPPEAT
    GuildAcceptAlliance     'ACEPALIA
    GuildOfferPeace         'PEACEOFF
    GuildOfferAlliance      'ALLIEOFF
    GuildAllianceDetails    'ALLIEDET
    GuildPeaceDetails       'PEACEDET
    GuildRequestJoinerInfo  'ENVCOMEN
    GuildAlliancePropList   'ENVALPRO
    GuildPeacePropList      'ENVPROPP
    GuildDeclareWar         'DECGUERR
    GuildNewWebsite         'NEWWEBSI
    GuildAcceptNewMember    'ACEPTARI
    GuildRejectNewMember    'RECHAZAR
    GuildKickMember         'ECHARCLA
    GuildUpdateNews         'ACTGNEWS
    GuildMemberInfo         '1HRINFO<
    GuildOpenElections      'ABREELEC
    GuildRequestMembership  'SOLICITUD
    GuildRequestDetails     'CLANDETAILS
    Online                  '/ONLINE
    Quit                    '/SALIR
    GuildLeave              '/SALIRCLAN
    RequestAccountState     '/BALANCE
    PetStand                '/QUIETO
    PetFollow               '/ACOMPAÑAR
    ReleasePet              '/LIBERAR
    TrainList               '/ENTRENAR
    Rest                    '/DESCANSAR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    Enlist                  '/ENLISTAR
    Information             '/INFORMACION
    Reward                  '/RECOMPENSA
    RequestMOTD             '/MOTD
    Uptime                  '/UPTIME
    PartyLeave              '/SALIRPARTY
    PartyCreate             '/CREARPARTY
    PartyJoin               '/PARTY
    Inquiry                 '/ENCUESTA ( with no params )
    GuildMessage            '/CMSG
    PartyMessage            '/PMSG
    CentinelReport          '/CENTINELA
    GuildOnline             '/ONLINECLAN
    PartyOnline             '/ONLINEPARTY
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    GMRequest               '/GM
    bugReport               '/_BUG
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    Punishments             '/PENAS
    ChangePassword          '/CONTRASEÑA
    Gamble                  '/APOSTAR
    InquiryVote             '/ENCUESTA ( with parameters )
    LeaveFaction            '/RETIRAR ( with no arguments )
    BankExtractGold         '/RETIRAR ( with arguments )
    BankDepositGold         '/DEPOSITAR
    Denounce                '/DENUNCIAR
    GuildFundate            '/FUNDARCLAN
    GuildFundation
    PartyKick               '/ECHARPARTY
    PartySetLeader          '/PARTYLIDER
    PartyAcceptMember       '/ACCEPTPARTY
    Ping                    '/PING
    RequestFormYesNo        ' GSZAO - Ventana SI o NO multi-uso
    RequestPartyForm
    ItemUpgrade
    GMCommands
    InitCrafting
    Home
    ShowGuildNews
    ShareNpc                '/COMPARTIRNPC
    StopSharingNpc          '/NOCOMPARTIRNPC
    Consultation
    
    Quest                   ' GSZAO - /QUEST
    QuestAccept             ' GSZAO
    QuestListRequest        ' GSZAO
    QuestDetailsRequest     ' GSZAO
    QuestAbandon            ' GSZAO
    
    MoveItem                'Drag and drop
    DropObj                 'Drop to pos.

End Enum

Public Enum eGMCommands
    GMMessage = 1           '/GMSG
    showName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    OnlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    Comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpMeToTarget          '/TELEPLOC
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    Invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestCharInfo         '/INFO
    RequestCharStats        '/STAT
    RequestCharGold         '/BAL
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Forgive                 '/PERDON
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    banChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    nickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    RainToggle              '/LLUVIA
    SetCharDescription      '/SETDESC
    ForceMIDIToMap          '/FORCEMIDIMAP
    ForceWAVEToMap          '/FORCEWAVMAP
    RoyalArmyMessage        '/REALMSG
    ChaosLegionMessage      '/CAOSMSG
    CitizenMessage          '/CIUMSG
    CriminalMessage         '/CRIMSG
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    AcceptRoyalCouncilMember '/ACEPTCONSE
    AcceptChaosCouncilMember '/ACEPTCONSECAOS
    ItemsInTheFloor         '/PISO
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    dumpIPTables            '/DUMPSECURITY
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    BanIP                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ChaosLegionKick         '/NOCAOS
    RoyalArmyKick           '/NOREAL
    ForceMIDIAll            '/FORCEMIDI
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    CreateNPC               '/ACC
    CreateNPCWithRespawn    '/RACC
    ImperialArmour          '/AI1 - 4
    ChaosArmour             '/AC1 - 4
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle '/HABILITAR
    TurnOffServer           '/APAGAR
    TurnCriminal            '/CONDEN
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/RAJARCLAN
    RequestCharMail         '/LASTEMAIL
    AlterPassword           '/APASS
    AlterMail               '/AEMAIL
    AlterName               '/ANAME
    ToggleCentinelActivated '/CENTINELAACTIVADO
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
    SaveMap                 '/GUARDARMAPA
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    
    ' 0.13.3
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    ChangeMapInfoStealNpc   '/MODMAPINFO ROBONPC
    ChangeMapInfoNoOcultar  '/MODMAPINFO OCULTARSINEFECTO
    ChangeMapInfoNoInvocar  '/MODMAPINFO INVOCARSINEFECTO
    
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    KickAllChars            '/ECHARTODOSPJS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    ResetAutoUpdate         '/AUTOUPDATE
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
    
    ' GSZAO
    AdminCargos             '/ADMIN <tipo> <accion> <nick> --  Tipos: D S C R   Accion: + - =
    VerHD                   '/VERHD USUARIO
    BanHD                   '/BANHD USUARIO
    UnBanHD                 '/UNBANHD NROHD
    SearchObj               '/BUSCAROBJ NOMBRE
    SearchNpc               '/BUSCARNPC NOMBRE
    LluviaDeORO             '/LLUVIADEORO
    
    ' 0.13.3
    CreatePretorianClan     '/CREARPRETORIANOS
    RemovePretorianClan     '/ELIMINARPRETORIANOS
    EnableDenounces         '/DENUNCIAS
    ShowDenouncesList       '/SHOW DENUNCIAS
    MapMessage              '/MAPMSG
    SetDialog               '/SETDIALOG
    Impersonate             '/IMPERSONAR
    Imitate                 '/MIMETIZAR
    RecordAdd
    RecordRemove
    RecordAddObs
    RecordListRequest
    RecordDetailsRequest
    ' 0.13.5
    AlterGuildName
    HigherAdminsMessage
End Enum

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_SEMIDIOSMSG
    FONTTYPE_SEMIDIOS
    FONTTYPE_CITIZEN
    FONTTYPE_CONSEJERO
    FONTTYPE_DIOS
    FONTTYPE_GOLD ' GSZAO
    FONTTYPE_OBJ ' GSZAO
    FONTTYPE_NPC_WARNING ' GSZAO
    FONTTYPE_NPC_PEACE ' GSZAO
End Enum

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_addGold
    eo_Vida
    eo_Poss
End Enum

Public Enum eCargos
    c_Rolmaster
    c_Consejero
    c_Semidios
    c_Dios
End Enum

Public Enum eAcciones
    a_Listar
    a_Agregar
    a_Quitar
End Enum

Public FontTypes(24) As tFont

''
' Initializes the fonts array

Public Sub InitFonts()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/07/2012 - ^[GS]^
'***************************************************
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .Red = 255
        .Green = 255
        .Blue = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .Red = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .Red = 32
        .Green = 51
        .Blue = 223
        .bold = 1
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .Red = 65
        .Green = 190
        .Blue = 156
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .Red = 65
        .Green = 190
        .Blue = 156
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .Red = 130
        .Green = 130
        .Blue = 130
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        .Red = 255
        .Green = 180
        .Blue = 250
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_VENENO).Green = 255
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .Red = 255
        .Green = 255
        .Blue = 255
        .bold = 1
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_SERVER).Green = 185
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        .Red = 228
        .Green = 199
        .Blue = 27
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
        .Red = 130
        .Green = 130
        .Blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .Red = 255
        .Green = 60
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
        .Green = 200
        .Blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
        .Red = 255
        .Green = 50
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
        .Green = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_SEMIDIOSMSG)
        .Red = 255
        .Green = 255
        .Blue = 255
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_SEMIDIOS)
        .Red = 30
        .Green = 255
        .Blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CITIZEN)
        .Blue = 200
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJERO)
        .Red = 30
        .Green = 150
        .Blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DIOS)
        .Red = 250
        .Green = 250
        .Blue = 150
        .bold = 1
    End With
    
   With FontTypes(FontTypeNames.FONTTYPE_GOLD) ' GSZAO
        .Red = 255
        .Green = 215
        .Blue = 0
        .bold = 1
    End With
    
   With FontTypes(FontTypeNames.FONTTYPE_OBJ) ' GSZAO
        .Red = 175
        .Green = 238
        .Blue = 238
        .bold = 1
    End With

   With FontTypes(FontTypeNames.FONTTYPE_NPC_WARNING) ' GSZAO
        .Red = 235
        .Green = 51
        .Blue = 51
        .bold = 1
    End With
    
   With FontTypes(FontTypeNames.FONTTYPE_NPC_PEACE)    ' GSZAO
        .Red = 51
        .Green = 235
        .Blue = 51
        .bold = 1
    End With

End Sub

''
' Handles incoming data.

Public Sub HandleIncomingData()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 18/03/2013 - ^[GS]^
'
'***************************************************
On Error Resume Next

    Dim Packet As Byte

    Packet = incomingData.PeekByte()
#If Testeo = 1 Then
    Debug.Print Now & " - HandleIncomingData " & Packet
#End If
    
    Select Case Packet
        Case ServerPacketID.Logged                  ' LOGGED
            Call HandleLogged
        
        Case ServerPacketID.InfoTorneo
            Call HandleInfoTorneo
            
        Case ServerPacketID.ClientConfig            ' CLIENTCFG - GSZ
            Call HandleClientConfig
        
        Case ServerPacketID.CreateParticleInChar    ' CPCHAR - GSZ
            Call HandleCreateParticleInChar
            
        Case ServerPacketID.RemoveDialogs           ' QTDL
            Call HandleRemoveDialogs
        
        Case ServerPacketID.RemoveCharDialog        ' QDL
            Call HandleRemoveCharDialog
        
        Case ServerPacketID.NavigateToggle          ' NAVEG
            Call HandleNavigateToggle
            
        Case ServerPacketID.CreateRenderValue       ' CDMG ' GSZAO
            Call HandleCreateRenderValue
           
        Case ServerPacketID.Disconnect              ' FINOK
            Call HandleDisconnect
        
        Case ServerPacketID.CommerceEnd             ' FINCOMOK
            Call HandleCommerceEnd
            
        Case ServerPacketID.CommerceChat
            Call HandleCommerceChat
        
        Case ServerPacketID.BankEnd                 ' FINBANOK
            Call HandleBankEnd
        
        Case ServerPacketID.CommerceInit            ' INITCOM
            Call HandleCommerceInit
        
        Case ServerPacketID.BankInit                ' INITBANCO
            Call HandleBankInit
        
        Case ServerPacketID.UserCommerceInit        ' INITCOMUSU
            Call HandleUserCommerceInit
        
        Case ServerPacketID.UserCommerceEnd         ' FINCOMUSUOK
            Call HandleUserCommerceEnd
            
        Case ServerPacketID.UserOfferConfirm
            Call HandleUserOfferConfirm
        
        Case ServerPacketID.ShowBlacksmithForm      ' SFH
            Call HandleShowBlacksmithForm
        
        Case ServerPacketID.ShowCarpenterForm       ' SFC
            Call HandleShowCarpenterForm
        
        Case ServerPacketID.UpdateSta               ' ASS
            Call HandleUpdateSta
        
        Case ServerPacketID.UpdateMana              ' ASM
            Call HandleUpdateMana
        
        Case ServerPacketID.UpdateHP                ' ASH
            Call HandleUpdateHP
        
        Case ServerPacketID.UpdateGold              ' ASG
            Call HandleUpdateGold
            
        Case ServerPacketID.UpdateBankGold
            Call HandleUpdateBankGold

        Case ServerPacketID.UpdateExp               ' ASE
            Call HandleUpdateExp
        
        Case ServerPacketID.ChangeMap               ' CM
            Call HandleChangeMap
        
        Case ServerPacketID.PosUpdate               ' PU
            Call HandlePosUpdate
        
        Case ServerPacketID.ChatOverHead            ' ||
            Call HandleChatOverHead
        
        Case ServerPacketID.ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
            Call HandleConsoleMessage
        
        Case ServerPacketID.GuildChat               ' |+
            Call HandleGuildChat
        
        Case ServerPacketID.ShowMessageBox          ' !!
            Call HandleShowMessageBox
        
        Case ServerPacketID.UserIndexInServer       ' IU
            Call HandleUserIndexInServer
        
        Case ServerPacketID.UserCharIndexInServer   ' IP
            Call HandleUserCharIndexInServer
        
        Case ServerPacketID.CharacterCreate         ' CC
            Call HandleCharacterCreate
        
        Case ServerPacketID.CharacterRemove         ' BP
            Call HandleCharacterRemove
        
        Case ServerPacketID.CharacterChangeNick
            Call HandleCharacterChangeNick
            
        Case ServerPacketID.CharacterMove           ' MP, +, * and _ '
            Call HandleCharacterMove
            
        Case ServerPacketID.ForceCharMove
            Call HandleForceCharMove
        
        Case ServerPacketID.CharacterChange         ' CP
            Call HandleCharacterChange
        
        Case ServerPacketID.ObjectCreate            ' HO
            Call HandleObjectCreate
        
        Case ServerPacketID.ObjectDelete            ' BO
            Call HandleObjectDelete
        
        Case ServerPacketID.BlockPosition           ' BQ
            Call HandleBlockPosition
        
        Case ServerPacketID.PlayMIDI                ' TM
            Call HandlePlayMIDI
        
        Case ServerPacketID.PlayWave                ' TW
            Call HandlePlayWave
        
        Case ServerPacketID.guildList               ' GL
            Call HandleGuildList
        
        Case ServerPacketID.AreaChanged             ' CA
            Call HandleAreaChanged
        
        Case ServerPacketID.PauseToggle             ' BKW
            Call HandlePauseToggle
        
        Case ServerPacketID.RainToggle              ' LLU
            Call HandleRainToggle
        
        Case ServerPacketID.CreateFX                ' CFX
            Call HandleCreateFX
        
        Case ServerPacketID.UpdateUserStats         ' EST
            Call HandleUpdateUserStats
        
        Case ServerPacketID.WorkRequestTarget       ' T01
            Call HandleWorkRequestTarget
        
        Case ServerPacketID.ChangeInventorySlot     ' CSI
            Call HandleChangeInventorySlot
        
        Case ServerPacketID.ChangeBankSlot          ' SBO
            Call HandleChangeBankSlot
        
        Case ServerPacketID.ChangeSpellSlot         ' SHS
            Call HandleChangeSpellSlot
        
        Case ServerPacketID.Atributes               ' ATR
            Call HandleAtributes
        
        Case ServerPacketID.BlacksmithWeapons       ' LAH
            Call HandleBlacksmithWeapons
        
        Case ServerPacketID.BlacksmithArmors        ' LAR
            Call HandleBlacksmithArmors
        
        Case ServerPacketID.CarpenterObjects        ' OBR
            Call HandleCarpenterObjects
        
        Case ServerPacketID.RestOK                  ' DOK
            Call HandleRestOK
        
        Case ServerPacketID.ErrorMsg                ' ERR
            Call HandleErrorMessage
        
        Case ServerPacketID.Blind                   ' CEGU
            Call HandleBlind
        
        Case ServerPacketID.Dumb                    ' DUMB
            Call HandleDumb
        
        Case ServerPacketID.ShowSignal              ' MCAR
            Call HandleShowSignal
        
        Case ServerPacketID.ChangeNPCInventorySlot  ' NPCI
            Call HandleChangeNPCInventorySlot
        
        Case ServerPacketID.UpdateHungerAndThirst   ' EHYS
            Call HandleUpdateHungerAndThirst
        
        Case ServerPacketID.Fame                    ' FAMA
            Call HandleFame
        
        Case ServerPacketID.MiniStats               ' MEST
            Call HandleMiniStats
        
        Case ServerPacketID.LevelUp                 ' SUNI
            Call HandleLevelUp
        
        Case ServerPacketID.AddForumMsg             ' FMSG
            Call HandleAddForumMessage
        
        Case ServerPacketID.ShowForumForm           ' MFOR
            Call HandleShowForumForm
        
        Case ServerPacketID.SetInvisible            ' NOVER
            Call HandleSetInvisible
        
        Case ServerPacketID.DiceRoll                ' DADOS
            Call HandleDiceRoll
        
        Case ServerPacketID.MeditateToggle          ' MEDOK
            Call HandleMeditateToggle
        
        Case ServerPacketID.BlindNoMore             ' NSEGUE
            Call HandleBlindNoMore
        
        Case ServerPacketID.DumbNoMore              ' NESTUP
            Call HandleDumbNoMore
        
        Case ServerPacketID.SendSkills              ' SKILLS
            Call HandleSendSkills
        
        Case ServerPacketID.TrainerCreatureList     ' LSTCRI
            Call HandleTrainerCreatureList
        
        Case ServerPacketID.guildNews               ' GUILDNE
            Call HandleGuildNews
        
        Case ServerPacketID.OfferDetails            ' PEACEDE and ALLIEDE
            Call HandleOfferDetails
        
        Case ServerPacketID.AlianceProposalsList    ' ALLIEPR
            Call HandleAlianceProposalsList
        
        Case ServerPacketID.PeaceProposalsList      ' PEACEPR
            Call HandlePeaceProposalsList
        
        Case ServerPacketID.CharacterInfo           ' CHRINFO
            Call HandleCharacterInfo
        
        Case ServerPacketID.GuildLeaderInfo         ' LEADERI
            Call HandleGuildLeaderInfo
        
        Case ServerPacketID.GuildDetails            ' CLANDET
            Call HandleGuildDetails
        
        Case ServerPacketID.ShowGuildFundationForm  ' SHOWFUN
            Call HandleShowGuildFundationForm
        
        Case ServerPacketID.ParalizeOK              ' PARADOK
            Call HandleParalizeOK
        
        Case ServerPacketID.ShowUserRequest         ' PETICIO
            Call HandleShowUserRequest
        
        Case ServerPacketID.TradeOK                 ' TRANSOK
            Call HandleTradeOK
        
        Case ServerPacketID.BankOK                  ' BANCOOK
            Call HandleBankOK
        
        Case ServerPacketID.ChangeUserTradeSlot     ' COMUSUINV
            Call HandleChangeUserTradeSlot
        
        Case ServerPacketID.Pong
            Call HandlePong
        
        Case ServerPacketID.UpdateTagAndStatus ' = 91
            Call HandleUpdateTagAndStatus
        
        Case ServerPacketID.GuildMemberInfo ' = 82
            Call HandleGuildMemberInfo
            
        
        
        '*******************
        'GM messages
        '*******************
        Case ServerPacketID.SpawnList               ' SPL
            Call HandleSpawnList
        
        Case ServerPacketID.ShowSOSForm             ' RSOS and MSOS
            Call HandleShowSOSForm
        
        Case ServerPacketID.ShowDenounces           ' 0.13.3
            Call HandleShowDenounces
            
        Case ServerPacketID.RecordDetails           ' 0.13.3
            Call HandleRecordDetails
            
        Case ServerPacketID.RecordList              ' 0.13.3
            Call HandleRecordList
        
        Case ServerPacketID.ShowMOTDEditionForm     ' ZMOTD
            Call HandleShowMOTDEditionForm
        
        Case ServerPacketID.ShowGMPanelForm         ' ABPANEL
            Call HandleShowGMPanelForm
        
        Case ServerPacketID.UserNameList            ' LISTUSU
            Call HandleUserNameList
            
        Case ServerPacketID.ShowGuildAlign
            Call HandleShowGuildAlign
        
        Case ServerPacketID.ShowPartyForm
            Call HandleShowPartyForm
        
        Case ServerPacketID.UpdateStrenghtAndDexterity
            Call HandleUpdateStrenghtAndDexterity
            
        Case ServerPacketID.UpdateStrenght
            Call HandleUpdateStrenght
            
        Case ServerPacketID.UpdateDexterity
            Call HandleUpdateDexterity
            
        Case ServerPacketID.AddSlots
            Call HandleAddSlots

        Case ServerPacketID.MultiMessage
            Call HandleMultiMessage
        
        Case ServerPacketID.StopWorking
            Call HandleStopWorking
            
        Case ServerPacketID.CancelOfferItem
            Call HandleCancelOfferItem
            
        Case ServerPacketID.QuestDetails        ' GSZAO
            Call HandleQuestDetails
       
        Case ServerPacketID.QuestListSend       ' GSZAO
            Call HandleQuestListSend
            
        Case ServerPacketID.FormYesNo           ' GSZAO
            Call HandleFormYesNo
            
        Case ServerPacketID.Mensajes            ' GSZAO
            Call HandleMensajes
            
        Case ServerPacketID.Online              ' GSZAO
            Call HandleOnline
        
        Case ServerPacketID.CreateParticle
            Call HandleCreateParticle
            
        Case ServerPacketID.UserDeath
            Call HandleDieAlocate
        Case Else
            'ERROR : Abort!
            Exit Sub
    End Select
    
    'Done with this packet, move on to next one
    If incomingData.length > 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then
        Err.Clear
        Call HandleIncomingData
    End If
End Sub

Public Sub HandleMultiMessage()
'***************************************************
'Author: Unknown
'Last Modification: 18/03/2013 - ^[GS]^
'***************************************************

#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleMultiMessage"
#End If


    Dim BodyPart As Byte
    Dim Daño As Integer
    
With incomingData
    Call .ReadByte
    
    Select Case .ReadByte
        Case eMessages.DontSeeAnything
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_NO_VES_NADA_INTERESANTE, 65, 190, 156, False, False, True)
        
        Case eMessages.NPCSwing
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, True)
        
        Case eMessages.NPCKillUser
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, True)
        
        Case eMessages.BlockedWithShieldUser
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
        
        Case eMessages.BlockedWithShieldOther
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
        
        Case eMessages.UserSwing
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, True)
        
        Case eMessages.SafeModeOn
            Call frmMain.ControlSM(eSMType.sSafemode, True)
        
        Case eMessages.SafeModeOff
            Call frmMain.ControlSM(eSMType.sSafemode, False)
        
        Case eMessages.ResuscitationSafeOff
            Call frmMain.ControlSM(eSMType.sResucitation, False)
         
        Case eMessages.ResuscitationSafeOn
            Call frmMain.ControlSM(eSMType.sResucitation, True)
        
        Case eMessages.NobilityLost
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, True)
        
        Case eMessages.CantUseWhileMeditating
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, True)
        
        Case eMessages.NPCHitUser
            Select Case incomingData.ReadByte()
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & CStr(incomingData.ReadInteger() & "!!"), 255, 0, 0, True, False, True)
            End Select
        
        Case eMessages.UserHitNPC
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & CStr(incomingData.ReadLong()) & MENSAJE_2, 255, 0, 0, True, False, True)
        
        Case eMessages.UserAttackedSwing
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & CharList(incomingData.ReadInteger()).Nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, True)
        
        Case eMessages.UserHittedByUser
            Dim AttackerName As String
            
            AttackerName = GetRawName(CharList(incomingData.ReadInteger()).Nombre)
            BodyPart = incomingData.ReadByte()
            Daño = incomingData.ReadInteger()
            
            Select Case BodyPart
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIBE_IMPACTO_CABEZA & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIBE_IMPACTO_BRAZO_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIBE_IMPACTO_BRAZO_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIBE_IMPACTO_PIERNA_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIBE_IMPACTO_PIERNA_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIBE_IMPACTO_TORSO & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
            End Select
        
        Case eMessages.UserHittedUser

            Dim VictimName As String
            
            VictimName = GetRawName(CharList(incomingData.ReadInteger()).Nombre)
            BodyPart = incomingData.ReadByte()
            Daño = incomingData.ReadInteger()
            
            Select Case BodyPart
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_CABEZA & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_TORSO & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
            End Select
        
        Case eMessages.WorkRequestTarget
            UsingSkill = incomingData.ReadByte()
            
            Call ChangeCursorMain(cur_Action)
            
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
                
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
                
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
                
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
                
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
                
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
                    
                ' GSZAO
                Case eAccionClick.Matrimonio
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_MATROMONIO, 100, 100, 120, 0, 0)
                ' GSZAO
                
                Case eAccionClick.Divorcio
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_DIVORCIO, 100, 100, 120, 0, 0)
                    
            End Select

        Case eMessages.HaveKilledUser ' 0.13.3
            Dim KilledUser As Integer
            Dim Exp As Long
            
            KilledUser = .ReadInteger
            Exp = .ReadLong
            
            Call ShowConsoleMsg(MENSAJE_HAS_MATADO_A & CharList(KilledUser).Nombre & MENSAJE_22, 255, 0, 0, True, False)
            Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXPE_1 & Exp & MENSAJE_HAS_GANADO_EXPE_2, 255, 0, 0, True, False)
            
            'Sacamos un screenshot si está activado el FragShooter:
            If ClientAOSetup.bKill And ClientAOSetup.bActive Then
                If Exp \ 2 > ClientAOSetup.byMurderedLevel Then
                    FragShooterNickname = CharList(KilledUser).Nombre
                    FragShooterKilledSomeone = True
                    FragShooterCapturePending = True
                End If
            End If
            
        Case eMessages.UserKill ' 0.13.3
            Dim KillerUser As Integer
            
            KillerUser = .ReadInteger
            
            Call ShowConsoleMsg(CharList(KillerUser).Nombre & MENSAJE_TE_HA_MATADO, 255, 0, 0, True, False)
            
            'Sacamos un screenshot si está activado el FragShooter:
            If ClientAOSetup.bDie And ClientAOSetup.bActive Then
                FragShooterNickname = CharList(KillerUser).Nombre
                FragShooterKilledSomeone = False
                FragShooterCapturePending = True
            End If
                
        Case eMessages.EarnExp ' 0.13.3
            'Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXPE_1 & .ReadLong & MENSAJE_HAS_GANADO_EXPE_2, 255, 0, 0, True, False)
        
        Case eMessages.GoHome ' 0.13.3
            Dim Distance As Byte
            Dim Hogar As String
            Dim tiempo As Integer
            Dim msg As String
            
            Distance = .ReadByte
            tiempo = .ReadInteger
            Hogar = .ReadASCIIString
            
            If tiempo >= 60 Then
                If tiempo Mod 60 = 0 Then
                    msg = tiempo / 60 & " minutos."
                Else
                    msg = CInt(tiempo \ 60) & " minutos y " & tiempo Mod 60 & " segundos."  'Agregado el CInt() asi el número no es con , [C4b3z0n - 09/28/2010]
                End If
            Else
                msg = tiempo & " segundos."
            End If
            
            Call ShowConsoleMsg("Te encuentras a " & Distance & " mapas de la " & Hogar & ", este viaje durará " & msg, 255, 0, 0, True)
            Traveling = True
        
        Case eMessages.FinishHome
            Call ShowConsoleMsg(MENSAJE_HOGAR, 255, 255, 255)
            Traveling = False
        
        Case eMessages.CancelGoHome
            Call ShowConsoleMsg(MENSAJE_HOGAR_CANCEL, 255, 0, 0, True)
            Traveling = False
    End Select
End With

End Sub



''
' Handles the Logged message.

Private Sub HandleLogged()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/08/2013 - ^[GS]^
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleLogged"
#End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    ' Variable initialization
    UserClase = incomingData.ReadByte
    EngineRun = True
    Nombres = True
    bRain = False
    bRangoReducido = False
    
    'Set connected state
    Call SetConnected
    
    If bShowTutorial Then frmTutorial.Show vbModeless
    
    Inventario.DrawInv
    Spells.RenderSpells
    
    'Show tip
    'If tipf = "1" And PrimeraVez Then
    '    Call CargarTip
    '    frmtip.Visible = True
    '    PrimeraVez = False
    'End If
End Sub

''
' Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveAllDialogs
End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check if the packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveDialog(incomingData.ReadInteger())
End Sub

''
' Handles the NavigateToggle message.

Private Sub HandleNavigateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserNavegando = Not UserNavegando
End Sub

Private Sub HandleCreateRenderValue()
'***************************************************
'Author: maTih.-
'Last Modification: 09/06/2012 - ^[GS]^
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleCreateRenderValue"
#End If
    
    With incomingData
         .ReadByte
         Call modRenderValue.Create(.ReadByte(), .ReadByte(), 0, .ReadInteger(), .ReadByte())
    End With
End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Close connection
    frmMain.Socket1.Disconnect
    
    ResetAllInfo ' 0.13.3

End Sub

''
' Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/04/2013 - ^[GS]^
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvComUsu = Nothing
    Set InvComNpc = Nothing
    
    If Comerciando = True Then
        'Hide form
        Unload frmComerciar
        'Reset vars
        Comerciando = False
    End If
    
End Sub

''
' Handles the BankEnd message.

Private Sub HandleBankEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvBanco(0) = Nothing
    Set InvBanco(1) = Nothing
    
    Unload frmBancoObj
    Comerciando = False
End Sub

''
' Handles the CommerceInit message.

Private Sub HandleCommerceInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvComUsu = New clsGraphicalInventory
    Set InvComNpc = New clsGraphicalInventory
    
    ' Initialize commerce inventories
    Call InvComUsu.Initialize(DirectD3D8, frmComerciar.picInvUser, Inventario.MaxObjs)
    Call InvComNpc.Initialize(DirectD3D8, frmComerciar.picInvNpc, MAX_NPC_INVENTORY_SLOTS)

    'Fill user inventory
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.OBJIndex(i) <> 0 Then
            With Inventario
                Call InvComUsu.SetItem(i, .OBJIndex(i), _
                .amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .Valor(i), .ItemName(i))
            End With
        End If
    Next i
    
    ' Fill Npc inventory
    For i = 1 To 50
        If NPCInventory(i).OBJIndex <> 0 Then
            With NPCInventory(i)
                Call InvComNpc.SetItem(i, .OBJIndex, _
                .amount, 0, .GrhIndex, _
                .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                .Valor, .Name)
            End With
        End If
    Next i
    
    'Set state and show form
    Comerciando = True
    frmComerciar.Show , frmMain
End Sub

''
' Handles the BankInit message.

Private Sub HandleBankInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    Dim i As Long
    Dim BankGold As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvBanco(0) = New clsGraphicalInventory
    Set InvBanco(1) = New clsGraphicalInventory
    
    BankGold = incomingData.ReadLong
    
    Call InvBanco(0).Initialize(DirectD3D8, frmBancoObj.PicBancoInv, MAX_BANCOINVENTORY_SLOTS)
    Call InvBanco(1).Initialize(DirectD3D8, frmBancoObj.PicInv, Inventario.MaxObjs)
    
    For i = 1 To Inventario.MaxObjs
        With Inventario
            Call InvBanco(1).SetItem(i, .OBJIndex(i), _
                .amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .Valor(i), .ItemName(i))
        End With
    Next i
    
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        With UserBancoInventory(i)
            Call InvBanco(0).SetItem(i, .OBJIndex, _
                .amount, .Equipped, .GrhIndex, _
                .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                .Valor, .Name)
        End With
    Next i
    
    'Set state and show form
    Comerciando = True
    
    frmBancoObj.lblUserGld.Caption = BankGold
    
    frmBancoObj.Show , frmMain
End Sub

''
' Handles the UserCommerceInit message.

Private Sub HandleUserCommerceInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 11/03/2012 - ^[GS]^
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    TradingUserName = incomingData.ReadASCIIString
    
    Set InvComUsu = New clsGraphicalInventory
    Set InvOfferComUsu(0) = New clsGraphicalInventory
    Set InvOfferComUsu(1) = New clsGraphicalInventory
    Set InvOroComUsu(0) = New clsGraphicalInventory
    Set InvOroComUsu(1) = New clsGraphicalInventory
    Set InvOroComUsu(2) = New clsGraphicalInventory

    ' Initialize commerce inventories
    Call InvComUsu.Initialize(DirectD3D, frmComerciarUsu.picInvComercio, Inventario.MaxObjs)
    Call InvOfferComUsu(0).Initialize(DirectD3D8, frmComerciarUsu.picInvOfertaProp, INV_OFFER_SLOTS)
    Call InvOfferComUsu(1).Initialize(DirectD3D8, frmComerciarUsu.picInvOfertaOtro, INV_OFFER_SLOTS)
    Call InvOroComUsu(0).Initialize(DirectD3D8, frmComerciarUsu.picInvOroProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    Call InvOroComUsu(1).Initialize(DirectD3D8, frmComerciarUsu.picInvOroOfertaProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    Call InvOroComUsu(2).Initialize(DirectD3D8, frmComerciarUsu.picInvOroOfertaOtro, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    
    'Fill user inventory
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.OBJIndex(i) <> 0 Then
            With Inventario
                Call InvComUsu.SetItem(i, .OBJIndex(i), _
                .amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .Valor(i), .ItemName(i))
            End With
        End If
    Next i

    ' Inventarios de oro
    Call InvOroComUsu(0).SetItem(1, ORO_INDEX, UserGLD, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")
    Call InvOroComUsu(1).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")
    Call InvOroComUsu(2).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")

    'Set state and show form
    Comerciando = True
    Call frmComerciarUsu.Show(vbModeless, frmMain)
    
End Sub

''
' Handles the UserCommerceEnd message.

Private Sub HandleUserCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/04/2013 - ^[GS]^
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvComUsu = Nothing
    Set InvOroComUsu(0) = Nothing
    Set InvOroComUsu(1) = Nothing
    Set InvOroComUsu(2) = Nothing
    Set InvOfferComUsu(0) = Nothing
    Set InvOfferComUsu(1) = Nothing
    
    'Destroy the form and reset the state
    If Comerciando = True Then
        Unload frmComerciarUsu
        Comerciando = False
    End If
    
End Sub

''
' Handles the UserOfferConfirm message.
Private Sub HandleUserOfferConfirm()
'***************************************************
'Author: ZaMa
'Last Modification: 27/07/2012 - ^[GS]^
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    With frmComerciarUsu
        ' Now he can accept the offer or reject it
        .HabilitarAceptarRechazar True
        
        .PrintCommerceMsg "¡" & TradingUserName & " ha confirmado su oferta!", FontTypeNames.FONTTYPE_CONSEJERO
    End With
    
End Sub

''
' Handles the ShowBlacksmithForm message.

Private Sub HandleShowBlacksmithForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/08/2014 - ^[GS]^
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If frmMain.MacroTrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftBlacksmith(MacroBltIndex)
    Else
        frmConstruirHerrero.Show , frmMain
        MirandoHerreria = True
    End If
End Sub

''
' Handles the ShowCarpenterForm message.

Private Sub HandleShowCarpenterForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/08/2014 - ^[GS]^
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If frmMain.MacroTrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftCarpenter(MacroBltIndex)
    Else
        frmConstruirCarp.Show , frmMain
        MirandoCarpinteria = True
    End If
End Sub

''
' Handles the NPCSwing message.

Private Sub HandleNPCSwing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, True)
End Sub

''
' Handles the NPCKillUser message.

Private Sub HandleNPCKillUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, True)
End Sub

''
' Handles the BlockedWithShieldUser message.

Private Sub HandleBlockedWithShieldUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
End Sub

''
' Handles the BlockedWithShieldOther message.

Private Sub HandleBlockedWithShieldOther()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
End Sub

''
' Handles the UserSwing message.

Private Sub HandleUserSwing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, True)
End Sub

''
' Handles the SafeModeOn message.

Private Sub HandleSafeModeOn()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleSafeModeOn"
#End If

    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sSafemode, True)
End Sub

''
' Handles the SafeModeOff message.

Private Sub HandleSafeModeOff()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleSafeModeOff"
#End If

    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sSafemode, False)
End Sub

''
' Handles the ResuscitationSafeOff message.

Private Sub HandleResuscitationSafeOff()
'***************************************************
'Author: Rapsodius
'Creation date: 10/10/07
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sResucitation, False)
End Sub

''
' Handles the ResuscitationSafeOn message.

Private Sub HandleResuscitationSafeOn()
'***************************************************
'Author: Rapsodius
'Creation date: 10/10/07
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sResucitation, True)
End Sub

''
' Handles the NobilityLost message.

Private Sub HandleNobilityLost()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, True)
End Sub

''
' Handles the CantUseWhileMeditating message.

Private Sub HandleCantUseWhileMeditating()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, True)
End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 25/08/2012 - ^[GS]^
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinSTA = incomingData.ReadInteger()
    
    frmMain.cStatEnergia.Value = UserMinSTA
End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 25/08/2012 - ^[GS]^
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinMAN = incomingData.ReadInteger()
    
    frmMain.cStatMana.Value = UserMinMAN
End Sub

''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 25/08/2012 - ^[GS]^
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinHP = incomingData.ReadInteger()
    
    frmMain.cStatVida.Value = UserMinHP
    
    'Is the user alive??
    If UserMinHP = 0 Then
        UserEstado = 1
        If frmMain.TrainingMacro Then Call frmMain.DesactivarMacroHechizos
        If frmMain.MacroTrabajo Then Call frmMain.DesactivarMacroTrabajo
    Else
        UserEstado = 0
    End If
End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateGold()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 02/05/2013 - ^[GS]^
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserGLD = incomingData.ReadLong()
    
    If UserGLD >= CLng(UserLvl) * 10000 And UserLvl > 12 Then 'Si el nivel es mayor de 12, es decir, no es newbie.
        'Changes color
        frmMain.lblOro.ForeColor = &H80FF&     'Orange
    Else
        'Changes color
        frmMain.lblOro.ForeColor = &HD7FF&     'Gold
    End If
    
    frmMain.lblOro.Value = UserGLD
End Sub

''
' Handles the UpdateBankGold message.

Private Sub HandleUpdateBankGold()
'***************************************************
'Autor: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmBancoObj.lblUserGld.Caption = incomingData.ReadLong
    
End Sub

''
' Handles the UpdateExp message.

Private Sub HandleUpdateExp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/09/2012 - ^[GS]^
'***************************************************
    'Check packet is complete
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserExp = incomingData.ReadLong()
    
    frmMain.cStatExp.Value = UserExp
End Sub

''
' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenghtAndDexterity()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserFuerza = incomingData.ReadByte
    UserAgilidad = incomingData.ReadByte
    frmMain.lblStrg.Caption = UserFuerza
    frmMain.lblDext.Caption = UserAgilidad
    frmMain.lblStrg.ForeColor = getStrenghtColor()
    frmMain.lblDext.ForeColor = getDexterityColor()
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenght()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************
    'Check packet is complete
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserFuerza = incomingData.ReadByte
    frmMain.lblStrg.Caption = UserFuerza
    frmMain.lblStrg.ForeColor = getStrenghtColor()
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateDexterity()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************
    'Check packet is complete
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserAgilidad = incomingData.ReadByte
    frmMain.lblDext.Caption = UserAgilidad
    frmMain.lblDext.ForeColor = getDexterityColor()
End Sub

''
' Handles the ChangeMap message.
Private Sub HandleChangeMap()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 25/07/2012 - ^[GS]^
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleChangeMap"
#End If
    
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMap = incomingData.ReadInteger()
    mapInfo.Pk = incomingData.ReadBoolean() ' GSZAO
    'Call incomingData.ReadInteger
           
    Call SwitchMap(UserMap)
    If bLluvia(UserMap) = 0 Then
        If bRain Then
            Call Audio.StopWave(RainBufferIndex)
            RainBufferIndex = 0
            frmMain.IsPlaying = PlayLoop.plNone
        End If
    End If

End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/09/2012 - ^[GS]^
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandlePosUpdate"
#End If
    
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Remove char from old position
    If MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex Then
        MapData(UserPos.X, UserPos.Y).CharIndex = 0
    End If
    
    'Set new pos
    UserPos.X = incomingData.ReadByte()
    UserPos.Y = incomingData.ReadByte()
    
    'Set char
    MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
    CharList(UserCharIndex).Pos = UserPos
    
    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2, True, False)
                
    'Update pos label
    Call UpdateUserPos
    
End Sub

''
' Handles the NPCHitUser message.

Private Sub HandleNPCHitUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Select Case incomingData.ReadByte()
        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & CStr(incomingData.ReadInteger() & "!!"), 255, 0, 0, True, False, True)
    End Select
End Sub

''
' Handles the UserHitNPC message.

Private Sub HandleUserHitNPC()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & CStr(incomingData.ReadLong()) & MENSAJE_2, 255, 0, 0, True, False, True)
End Sub

''
' Handles the UserAttackedSwing message.

Private Sub HandleUserAttackedSwing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & CharList(incomingData.ReadInteger()).Nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, True)
End Sub

''
' Handles the UserHittingByUser message.

Private Sub HandleUserHittedByUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim attacker As String
    
    attacker = CharList(incomingData.ReadInteger()).Nombre
    
    Select Case incomingData.ReadByte
        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIBE_IMPACTO_CABEZA & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIBE_IMPACTO_BRAZO_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIBE_IMPACTO_BRAZO_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIBE_IMPACTO_PIERNA_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIBE_IMPACTO_PIERNA_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIBE_IMPACTO_TORSO & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
    End Select
End Sub

''
' Handles the UserHittedUser message.

Private Sub HandleUserHittedUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Victim As String
    
    Victim = CharList(incomingData.ReadInteger()).Nombre
    
    Select Case incomingData.ReadByte
        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_CABEZA & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_TORSO & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
    End Select
End Sub

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleChatOverHead"
#End If
    
    If incomingData.length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Chat As String
    Dim CharIndex As Integer
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    
    Chat = Buffer.ReadASCIIString()
    CharIndex = Buffer.ReadInteger()
    
    r = Buffer.ReadByte()
    g = Buffer.ReadByte()
    b = Buffer.ReadByte()
    
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If CharList(CharIndex).active Then _
        Call Dialogos.CreateDialog(Trim$(Chat), CharIndex, RGB(r, g, b))
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleConsoleMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleConsoleMessage"
#End If
    
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Chat As String
    Dim FontIndex As Integer
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    
    Chat = Buffer.ReadASCIIString()
    FontIndex = Buffer.ReadByte()

    If InStr(1, Chat, "~") Then
        str = ReadField(2, Chat, 126)
            If Val(str) > 255 Then
                r = 255
            Else
                r = Val(str)
            End If
            
            str = ReadField(3, Chat, 126)
            If Val(str) > 255 Then
                g = 255
            Else
                g = Val(str)
            End If
            
            str = ReadField(4, Chat, 126)
            If Val(str) > 255 Then
                b = 255
            Else
                b = Val(str)
            End If
            
        Call AddtoRichTextBox(frmMain.RecTxt, Left$(Chat, InStr(1, Chat, "~") - 1), r, g, b, Val(ReadField(5, Chat, 126)) <> 0, Val(ReadField(6, Chat, 126)) <> 0)
    Else
        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, Chat, .Red, .Green, .Blue, .bold, .italic)
        End With
        
        ' Para no perder el foco cuando chatea por party
        If FontIndex = FontTypeNames.FONTTYPE_PARTY Then
            If MirandoParty Then frmParty.SendTxt.SetFocus
        End If
    End If
'    Call checkText(chat)
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the GuildChat message.

Private Sub HandleGuildChat()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Chat As String
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    Dim tmp As Integer
    Dim Cont As Integer
    
    
    Chat = Buffer.ReadASCIIString()
    
    If Not DialogosClanes.Activo Then
        If InStr(1, Chat, "~") Then
            str = ReadField(2, Chat, 126)
            If Val(str) > 255 Then
                r = 255
            Else
                r = Val(str)
            End If
            
            str = ReadField(3, Chat, 126)
            If Val(str) > 255 Then
                g = 255
            Else
                g = Val(str)
            End If
            
            str = ReadField(4, Chat, 126)
            If Val(str) > 255 Then
                b = 255
            Else
                b = Val(str)
            End If
            
            Call AddtoRichTextBox(frmMain.RecTxt, Left$(Chat, InStr(1, Chat, "~") - 1), r, g, b, Val(ReadField(5, Chat, 126)) <> 0, Val(ReadField(6, Chat, 126)) <> 0)
        Else
            With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
                Call AddtoRichTextBox(frmMain.RecTxt, Chat, .Red, .Green, .Blue, .bold, .italic)
            End With
        End If
    Else
        Call DialogosClanes.PushBackText(ReadField(1, Chat, 126))
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleCommerceChat()
'***************************************************
'Author: ZaMa
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Chat As String
    Dim FontIndex As Integer
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    
    Chat = Buffer.ReadASCIIString()
    FontIndex = Buffer.ReadByte()
    
    If InStr(1, Chat, "~") Then
        str = ReadField(2, Chat, 126)
            If Val(str) > 255 Then
                r = 255
            Else
                r = Val(str)
            End If
            
            str = ReadField(3, Chat, 126)
            If Val(str) > 255 Then
                g = 255
            Else
                g = Val(str)
            End If
            
            str = ReadField(4, Chat, 126)
            If Val(str) > 255 Then
                b = 255
            Else
                b = Val(str)
            End If
            
        Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, Left$(Chat, InStr(1, Chat, "~") - 1), r, g, b, Val(ReadField(5, Chat, 126)) <> 0, Val(ReadField(6, Chat, 126)) <> 0)
    Else
        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, Chat, .Red, .Green, .Blue, .bold, .italic)
        End With
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    frmMensaje.msg.Caption = Buffer.ReadASCIIString()
    frmMensaje.Show
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleUserIndexInServer"
#End If
    
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserIndex = incomingData.ReadInteger()
End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/09/2012 - ^[GS]^
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleUserCharIndexInServer"
#End If

    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCharIndex = incomingData.ReadInteger()
    UserPos = CharList(UserCharIndex).Pos
    
    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2, True, False)

    Call UpdateUserPos
End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 30/10/2012 - ^[GS]^
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleCharacterCreate"
#End If

    If incomingData.length < 24 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As E_Heading
    Dim X As Byte
    Dim Y As Byte
    Dim Weapon As Integer
    Dim Shield As Integer
    Dim Helmet As Integer
    Dim NickColor As Byte
    Dim Privs As Integer
    
    CharIndex = Buffer.ReadInteger()
    Body = Buffer.ReadInteger()
    Head = Buffer.ReadInteger()
    Heading = Buffer.ReadByte()
    X = Buffer.ReadByte()
    Y = Buffer.ReadByte()
    Weapon = Buffer.ReadInteger()
    Shield = Buffer.ReadInteger()
    Helmet = Buffer.ReadInteger()
    
    With CharList(CharIndex)
        Call SetCharacterFx(CharIndex, Buffer.ReadInteger(), Buffer.ReadInteger())
        
        .Nombre = Buffer.ReadASCIIString()
        NickColor = Buffer.ReadByte()
        
        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0
        End If
        
        ' GSZAO
        If (NickColor And eNickColor.ieNewbie) <> 0 Then
            .Newbie = 1
        Else
            .Newbie = 0
        End If
        ' GSZAO
        
        ' GSZAO
        If (NickColor And eNickColor.ieMuerto) = True Then
            .muerto = True
        Else
            .muerto = False
        End If
        ' GSZAO
        
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        
        Privs = Buffer.ReadByte()
        
        If Privs <> 0 Then
            'If the player belongs to a council AND is an admin, only whos as an admin
            If (Privs And PlayerType.ChaosCouncil) <> 0 And (Privs And PlayerType.User) = 0 Then
                Privs = Privs Xor PlayerType.ChaosCouncil
            End If
            
            If (Privs And PlayerType.RoyalCouncil) <> 0 And (Privs And PlayerType.User) = 0 Then
                Privs = Privs Xor PlayerType.RoyalCouncil
            End If
            
            'If the player is a RM, ignore other flags
            If Privs And PlayerType.RoleMaster Then
                Privs = PlayerType.RoleMaster
            End If
            
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .priv = Log(Privs) / Log(2)
        Else
            .priv = 0
        End If
        
        .bType = Buffer.ReadByte() ' GSZAO
    End With
    
    Call MakeChar(CharIndex, Body, Head, Heading, X, Y, Weapon, Shield, Helmet)
    
    Call RefreshAllChars
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandleCharacterChangeNick()
'***************************************************
'Author: Budi
'Last Modification: 07/23/09
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleCharacterChangeNick"
#End If

    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet id
    Call incomingData.ReadByte
    Dim CharIndex As Integer
    CharIndex = incomingData.ReadInteger
    CharList(CharIndex).Nombre = incomingData.ReadASCIIString
    
End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleCharacterRemove"
#End If

    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    Call EraseChar(CharIndex)
    Call RefreshAllChars
End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleCharacterMove"
#End If

    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim X As Byte
    Dim Y As Byte
    
    CharIndex = incomingData.ReadInteger()
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    With CharList(CharIndex)
        If .FXIndex >= 40 And .FXIndex <= 49 Then   'If it's meditating, we remove the FX
            .FXIndex = 0
        End If
        
        ' Play steps sounds if the user is not an admin of any kind
        If .priv <> 1 And .priv <> 2 And .priv <> 3 And .priv <> 5 And .priv <> 25 Then
            Call DoPasosFx(CharIndex)
        End If
    End With
    
    Call MoveCharbyPos(CharIndex, X, Y)
    
    Call RefreshAllChars
End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleForceCharMove"
#End If

    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Direccion As Byte
    
    Direccion = incomingData.ReadByte()

    Call MoveCharbyHead(UserCharIndex, Direccion)
    Call MoveScreen(Direccion)
    
    Call RefreshAllChars
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleCharacterChange"
#End If

    If incomingData.length < 18 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim tempint As Integer
    Dim headIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    With CharList(CharIndex)
        tempint = incomingData.ReadInteger()
        
        If tempint < LBound(BodyData()) Or tempint > UBound(BodyData()) Then
            .Body = BodyData(0)
            .iBody = 0
        Else
            .Body = BodyData(tempint)
            .iBody = tempint
        End If
        
        headIndex = incomingData.ReadInteger()
        
        If headIndex < LBound(HeadData()) Or headIndex > UBound(HeadData()) Then
            .Head = HeadData(0)
            .iHead = 0
        Else
            .Head = HeadData(headIndex)
            .iHead = headIndex
        End If
        
        .muerto = (headIndex = CASPER_HEAD)
        .Heading = incomingData.ReadByte()
        
        tempint = incomingData.ReadInteger()
        If tempint <> 0 Then .Arma = WeaponAnimData(tempint)
        
        tempint = incomingData.ReadInteger()
        If tempint <> 0 Then .Escudo = ShieldAnimData(tempint)
        
        tempint = incomingData.ReadInteger()
        If tempint <> 0 Then .Casco = CascoAnimData(tempint)
        
        Call SetCharacterFx(CharIndex, incomingData.ReadInteger(), incomingData.ReadInteger())
    End With
    
    Call RefreshAllChars
End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleObjectCreate"
#End If
    
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    MapData(X, Y).ObjGrh.GrhIndex = incomingData.ReadInteger()
    
    Call InitGrh(MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex)
End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    MapData(X, Y).ObjGrh.GrhIndex = 0
End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    If incomingData.ReadBoolean() Then
        MapData(X, Y).Blocked = 1
    Else
        MapData(X, Y).Blocked = 0
    End If
End Sub

''
' Handles the PlayMIDI message.

Private Sub HandlePlayMIDI()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandlePlayMIDI"
#End If

    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim currentMidi As Integer
    Dim Loops As Integer
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    currentMidi = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    If currentMidi Then
        If currentMidi > MP3_INITIAL_INDEX Then
            Call Audio.MusicMP3Play(ClientConfigInit.DirMultimedia & "\MP3\" & currentMidi & ".mp3") ' GSZAO
        Else
            Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", Loops)
        End If
    End If
    
End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Last Modified by: Rapsodius
'Added support for 3D Sounds.
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandlePlayWave"
#End If

    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
        
    Dim wave As Byte
    Dim srcX As Byte
    Dim srcY As Byte
    
    wave = incomingData.ReadByte()
    srcX = incomingData.ReadByte()
    srcY = incomingData.ReadByte()
        
    Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)
End Sub

''
' Handles the GuildList message.

Private Sub HandleGuildList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleGuildList"
#End If
    
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmGuildAdm
        'Clear guild's list
        .guildslist.Clear
        
        GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        Dim i As Long
        For i = 0 To UBound(GuildNames())
            Call .guildslist.AddItem(GuildNames(i))
        Next i
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
        
        .Show vbModeless, frmMain
    End With
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the AreaChanged message.

Private Sub HandleAreaChanged()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleAreaChanged"
#End If
    
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
        
    Call CambioDeArea(X, Y)
End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    pausa = Not pausa
End Sub

''
' Handles the RainToggle message.

Private Sub HandleRainToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 19/10/2012 - ^[GS]^
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
    
    bTecho = (MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2)
            
    If bRain Then
        If bLluvia(UserMap) Then
            'Stop playing the rain sound
            Call Audio.StopWave(RainBufferIndex)
            RainBufferIndex = 0
            If bTecho Then
                Call Audio.PlayWave("lluviainend.wav", 0, 0, LoopStyle.Disabled)
            Else
                Call Audio.PlayWave("lluviaoutend.wav", 0, 0, LoopStyle.Disabled)
            End If
            frmMain.IsPlaying = PlayLoop.plNone
        End If
    End If
    
    bRain = incomingData.ReadBoolean() ' GSZAO
    
    'bRain = Not bRain
End Sub

''
' Handles the CreateFX message.

Private Sub HandleCreateFX()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleCreateFX"
#End If
    
    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim fX As Integer
    Dim Loops As Integer
    
    CharIndex = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    Call SetCharacterFx(CharIndex, fX, Loops)
End Sub

''
' Handles the UpdateUserStats message.

Private Sub HandleUpdateUserStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 02/05/2013 - ^[GS]^
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleUpdateUserStats"
#End If
    
    If incomingData.length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMaxHP = incomingData.ReadInteger()
    UserMinHP = incomingData.ReadInteger()
    UserMaxMAN = incomingData.ReadInteger()
    UserMinMAN = incomingData.ReadInteger()
    UserMaxSTA = incomingData.ReadInteger()
    UserMinSTA = incomingData.ReadInteger()
    UserGLD = incomingData.ReadLong()
    UserLvl = incomingData.ReadByte()
    UserPasarNivel = incomingData.ReadLong()
    UserExp = incomingData.ReadLong()
    
    frmMain.cStatExp.Max = UserPasarNivel
    frmMain.cStatExp.Value = UserExp
        
    frmMain.lblOro.Value = UserGLD
    frmMain.lblLvl.Caption = UserLvl
    
    'Stats
    frmMain.cStatVida.Max = UserMaxHP
    frmMain.cStatMana.Max = UserMaxMAN
    frmMain.cStatEnergia.Max = UserMaxSTA
    
    frmMain.cStatVida.Value = UserMinHP
    frmMain.cStatMana.Value = UserMinMAN
    frmMain.cStatEnergia.Value = UserMinSTA
    
    'Spells.RenderSpells
    
    If UserMinHP = 0 Then
        UserEstado = 1
        If frmMain.TrainingMacro Then Call frmMain.DesactivarMacroHechizos
        If frmMain.MacroTrabajo Then Call frmMain.DesactivarMacroTrabajo
    Else
        UserEstado = 0
    End If
    
    If UserGLD >= CLng(UserLvl) * 10000 And UserLvl > 12 Then 'Si el nivel es mayor de 12, es decir, no es newbie.
        'Changes color
        frmMain.lblOro.ForeColor = &H80FF&     'Orange
    Else
        'Changes color
        frmMain.lblOro.ForeColor = &HD7FF&     'Gold
    End If
End Sub

''
' Handles the WorkRequestTarget message.

Private Sub HandleWorkRequestTarget()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 18/03/2013 - ^[GS]^
'
'***************************************************
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UsingSkill = incomingData.ReadByte()

    Call ChangeCursorMain(cur_Action)
    
    Select Case UsingSkill
        Case Magia
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
        Case Pesca
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
        Case Robar
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
        Case Talar
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
        Case Mineria
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
        Case FundirMetal
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
        Case Proyectiles
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
        ' GSZAO
        Case eAccionClick.Matrimonio
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_MATROMONIO, 100, 100, 120, 0, 0)
        ' GSZAO
        
        Case eAccionClick.Divorcio
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_DIVORCIO, 100, 100, 120, 0, 0)
            
    End Select
End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleChangeInventorySlot"
#End If
    
    If incomingData.length < 22 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    Dim OBJIndex As Integer
    Dim Name As String
    Dim amount As Integer
    Dim Equipped As Boolean
    Dim GrhIndex As Integer
    Dim OBJType As Byte
    Dim MaxHit As Integer
    Dim MinHit As Integer
    Dim MaxDef As Integer
    Dim MinDef As Integer
    Dim Value As Single
    
    slot = Buffer.ReadByte()
    OBJIndex = Buffer.ReadInteger()
    Name = Buffer.ReadASCIIString()
    amount = Buffer.ReadInteger()
    Equipped = Buffer.ReadBoolean()
    GrhIndex = Buffer.ReadInteger()
    OBJType = Buffer.ReadByte()
    MaxHit = Buffer.ReadInteger()
    MinHit = Buffer.ReadInteger()
    MaxDef = Buffer.ReadInteger()
    MinDef = Buffer.ReadInteger
    Value = Buffer.ReadSingle()
    
    If Equipped Then
        Select Case OBJType
            Case eObjType.otWeapon
                frmMain.lblWeapon = MinHit & "/" & MaxHit
                UserWeaponEqpSlot = slot
            Case eObjType.otArmadura
                frmMain.lblArmor = MinDef & "/" & MaxDef
                UserArmourEqpSlot = slot
            Case eObjType.otescudo
                frmMain.lblShielder = MinDef & "/" & MaxDef
                UserHelmEqpSlot = slot
            Case eObjType.otcasco
                frmMain.lblHelm = MinDef & "/" & MaxDef
                UserShieldEqpSlot = slot
        End Select
    Else
        Select Case slot
            Case UserWeaponEqpSlot
                frmMain.lblWeapon = "0/0"
                UserWeaponEqpSlot = 0
            Case UserArmourEqpSlot
                frmMain.lblArmor = "0/0"
                UserArmourEqpSlot = 0
            Case UserHelmEqpSlot
                frmMain.lblShielder = "0/0"
                UserHelmEqpSlot = 0
            Case UserShieldEqpSlot
                frmMain.lblHelm = "0/0"
                UserShieldEqpSlot = 0
        End Select
    End If
    
    Call Inventario.SetItem(slot, OBJIndex, amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, Value, Name)

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

' Handles the AddSlots message.
Private Sub HandleAddSlots()
'***************************************************
'Author: Budi
'Last Modification: 12/01/09
'
'***************************************************

    Call incomingData.ReadByte
    
    MaxInventorySlots = incomingData.ReadByte
End Sub

Private Sub HandleMensajes()
'***************************************************
'Author: TwIsT (GSZAO)
'Last Modification: 19/03/2013 - ^[GS]^
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleMensajes"
#End If

If incomingData.length < 3 Then
    Err.Raise incomingData.NotEnoughDataErrCode
    Exit Sub
End If

Call incomingData.ReadByte
Dim M As Integer
M = incomingData.ReadInteger
Select Case M 'By TwIsT
   Case Is = eMensajes.Mensaje001 ' "Comercio cancelado por el otro usuario."
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("Comercio cancelado por el otro usuario.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje002 ' "Has terminado de descansar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has terminado de descansar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje003 ' "¡¡¡Estás obstruyendo la vía pública, muévete o serás encarcelado!!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡¡Estás obstruyendo la vía pública, muévete o serás encarcelado!!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje004 ' "Has sido expulsado del clan. ¡El clan ha sumado un punto de antifacción!"
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("Has sido expulsado del clan. ¡El clan ha sumado un punto de antifacción!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje005 ' "¡¡Estás muerto!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje006 ' "Estás demasiado lejos del vendedor."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Estás demasiado lejos del vendedor.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje007 ' "El sacerdote no puede curarte debido a que estás demasiado lejos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El sacerdote no puede curarte debido a que estás demasiado lejos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje008 ' "Estas demasiado lejos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Estas demasiado lejos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje009 ' "La puerta esta cerrada con llave."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("La puerta esta cerrada con llave.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje010 ' "La puerta está cerrada con llave."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("La puerta está cerrada con llave.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje011 ' "Estás demasiado lejos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Estás demasiado lejos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje012 ' "No puedes hacer fogatas en zona segura."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes hacer fogatas en zona segura.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje013 ' "Has prendido la fogata."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has prendido la fogata.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje014 ' "La ley impide realizar fogatas en las ciudades."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("La ley impide realizar fogatas en las ciudades.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje015 ' "No has podido hacer fuego."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No has podido hacer fuego.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje016 ' "¡Has sido liberado!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Has sido liberado!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje017 ' "El usuario no está online."
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("El usuario no está online.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje018 ' "No puedes banear a al alguien de mayor jerarquía."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes banear a al alguien de mayor jerarquía.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje019 ' "El personaje ya se encuentra baneado."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El personaje ya se encuentra baneado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje020 ' "La mascota no atacará a ciudadanos si eres miembro del ejército real o tienes el seguro activado."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("La mascota no atacará a ciudadanos si eres miembro del ejército real o tienes el seguro activado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje021 ' "No tienes suficiente dinero."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes suficiente dinero.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje022 ' "No puedes cargar mas objetos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes cargar mas objetos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje023 ' "Lo siento, no estoy interesado en este tipo de objetos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Lo siento, no estoy interesado en este tipo de objetos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje024 ' "Las armaduras del ejército real sólo pueden ser vendidas a los sastres reales."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Las armaduras del ejército real sólo pueden ser vendidas a los sastres reales.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje025 ' "Las armaduras de la legión oscura sólo pueden ser vendidas a los sastres del demonio."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Las armaduras de la legión oscura sólo pueden ser vendidas a los sastres del demonio.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje026 ' "No puedes vender ítems."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("No puedes vender ítems.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje027 ' "Mapa exclusivo para newbies."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Mapa exclusivo para newbies.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje028 ' "Mapa exclusivo para miembros del ejército real."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Mapa exclusivo para miembros del ejército real.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje029 ' "Mapa exclusivo para miembros de la legión oscura."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Mapa exclusivo para miembros de la legión oscura.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje030 ' "Solo se permite entrar al mapa si eres miembro de alguna facción."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Solo se permite entrar al mapa si eres miembro de alguna facción.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje031 ' "Comercio cancelado. El otro usuario se ha desconectado."
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("Comercio cancelado. El otro usuario se ha desconectado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje032 ' "¡¡Estás muriendo de frío, abrigate o morirás!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡Estás muriendo de frío, abrigate o morirás!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje033 ' "¡¡Has muerto de frío!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡Has muerto de frío!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje034 ' "¡¡Quitate de la lava, te estás quemando!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡Quitate de la lava, te estás quemando!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje035 ' "¡¡Has muerto quemado!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡Has muerto quemado!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje036 ' "Recuperas tu apariencia normal."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Recuperas tu apariencia normal.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje037 ' "Has vuelto a ser visible."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has vuelto a ser visible.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje038 ' "Estás envenenado, si no te curas morirás."
      With FontTypes(FontTypeNames.FONTTYPE_VENENO)
         Call ShowConsoleMsg("Estás envenenado, si no te curas morirás.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje039 ' "Has sanado."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has sanado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje040 ' "Gracias por jugar Argentum Online"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Gracias por jugar Argentum Online", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje041 ' "No puedes tirar objetos newbie."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes tirar objetos newbie.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje042 ' "No hay espacio en el piso."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay espacio en el piso.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje043 ' "¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU BARCA!"
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU BARCA!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje044 ' "No puedes cargar más objetos."
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("No puedes cargar más objetos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje045 ' "No hay nada aquí."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay nada aquí.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje046 ' "Sólo los newbies pueden usar este objeto."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Sólo los newbies pueden usar este objeto.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje047 ' "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. "
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje048 ' "Sólo los newbies pueden usar estos objetos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Sólo los newbies pueden usar estos objetos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje049 ' "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje050 ' "Antes de usar la herramienta deberías equipartela."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Antes de usar la herramienta deberías equipartela.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje051 ' "Debes tener equipada la herramienta para trabajar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Debes tener equipada la herramienta para trabajar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje052 ' "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo. "
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo. ", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje053 ' "¡¡Debes esperar unos momentos para tomar otra poción!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡Debes esperar unos momentos para tomar otra poción!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje054 ' "Te has curado del envenenamiento."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Te has curado del envenenamiento.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje055 ' "Sientes un gran mareo y pierdes el conocimiento."
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Sientes un gran mareo y pierdes el conocimiento.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje056 ' "Has abierto la puerta."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has abierto la puerta.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje057 ' "La llave no sirve."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("La llave no sirve.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje058 ' "Has cerrado con llave la puerta."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has cerrado con llave la puerta.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje059 ' "No está cerrada."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No está cerrada.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje060 ' "No hay agua allí."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay agua allí.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje061 ' "Estás demasiado hambriento y sediento."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Estás demasiado hambriento y sediento.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje062 ' "No tienes conocimientos de las Artes Arcanas."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes conocimientos de las Artes Arcanas.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje063 ' "No hay peligro aquí. Es zona segura."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay peligro aquí. Es zona segura.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje064 ' "Sólo miembros del ejército real pueden usar este cuerno."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Sólo miembros del ejército real pueden usar este cuerno.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje065 ' "Sólo miembros de la legión oscura pueden usar este cuerno."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Sólo miembros de la legión oscura pueden usar este cuerno.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje066 ' "Para recorrer los mares debes ser nivel 25 o superior."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Para recorrer los mares debes ser nivel 25 o superior.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje067 ' "Para recorrer los mares debes ser nivel 20 o superior."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Para recorrer los mares debes ser nivel 20 o superior.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje068 ' "¡Debes aproximarte al agua para usar el barco!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Debes aproximarte al agua para usar el barco!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje069 ' "Tu carisma y liderazgo no son suficientes para liderar una party."
      With FontTypes(FontTypeNames.FONTTYPE_PARTY)
         Call ShowConsoleMsg("Tu carisma y liderazgo no son suficientes para liderar una party.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje070 ' "Por el momento no se pueden crear más parties."
      With FontTypes(FontTypeNames.FONTTYPE_PARTY)
         Call ShowConsoleMsg("Por el momento no se pueden crear más parties.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje071 ' "La party está llena, no puedes entrar."
      With FontTypes(FontTypeNames.FONTTYPE_PARTY)
         Call ShowConsoleMsg("La party está llena, no puedes entrar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje072 ' "¡Has formado una party!"
      With FontTypes(FontTypeNames.FONTTYPE_PARTY)
         Call ShowConsoleMsg("¡Has formado una party!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje073 ' "No puedes hacerte líder."
      With FontTypes(FontTypeNames.FONTTYPE_PARTY)
         Call ShowConsoleMsg("No puedes hacerte líder.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje074 ' "¡Te has convertido en líder de la party!"
      With FontTypes(FontTypeNames.FONTTYPE_PARTY)
         Call ShowConsoleMsg("¡Te has convertido en líder de la party!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje075 ' "No tienes suficientes puntos de liderazgo para liderar una party."
      With FontTypes(FontTypeNames.FONTTYPE_PARTY)
         Call ShowConsoleMsg("No tienes suficientes puntos de liderazgo para liderar una party.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje076 ' "Ya perteneces a una party."
      With FontTypes(FontTypeNames.FONTTYPE_PARTY)
         Call ShowConsoleMsg("Ya perteneces a una party.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje077 ' "Ya perteneces a una party, escribe /SALIRPARTY para abandonarla"
      With FontTypes(FontTypeNames.FONTTYPE_PARTY)
         Call ShowConsoleMsg("Ya perteneces a una party, escribe /SALIRPARTY para abandonarla", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje078 ' "El fundador decidirá si te acepta en la party."
      With FontTypes(FontTypeNames.FONTTYPE_PARTY)
         Call ShowConsoleMsg("El fundador decidirá si te acepta en la party.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje079 ' "Para ingresar a una party debes hacer click sobre el fundador y luego escribir /PARTY"
      With FontTypes(FontTypeNames.FONTTYPE_PARTY)
         Call ShowConsoleMsg("Para ingresar a una party debes hacer click sobre el fundador y luego escribir /PARTY", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje080 ' "No eres miembro de ninguna party."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No eres miembro de ninguna party.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje081 ' "¡No eres el líder de tu party!"
      With FontTypes(FontTypeNames.FONTTYPE_PARTY)
         Call ShowConsoleMsg("¡No eres el líder de tu party!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje082 ' "¡Está muerto, no puedes aceptar miembros en ese estado!"
      With FontTypes(FontTypeNames.FONTTYPE_PARTY)
         Call ShowConsoleMsg("¡Está muerto, no puedes aceptar miembros en ese estado!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje083 ' "¡No se ha hecho el cambio de mando!"
      With FontTypes(FontTypeNames.FONTTYPE_PARTY)
         Call ShowConsoleMsg("¡No se ha hecho el cambio de mando!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje084 ' "¡Está muerto!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Está muerto!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje085 ' "No podés tener mas objetos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No podés tener mas objetos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje086 ' "No tienes mas espacio en el banco!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes mas espacio en el banco!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje087 ' "El banco no puede cargar tantos objetos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El banco no puede cargar tantos objetos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje088 ' "El centinela intenta llamar tu atención. ¡Respóndele rápido!"
      With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
         Call ShowConsoleMsg("El centinela intenta llamar tu atención. ¡Respóndele rápido!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje089 ' "El centinela intenta llamar tu atención. ¡Respondele rápido!"
      With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
         Call ShowConsoleMsg("El centinela intenta llamar tu atención. ¡Respondele rápido!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje090 ' "¡¡¡Has sido expulsado del ejército real!!!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("¡¡¡Has sido expulsado del ejército real!!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje091 ' "¡¡¡Te has retirado del ejército real!!!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("¡¡¡Te has retirado del ejército real!!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje092 ' "¡¡¡Has sido expulsado de la Legión Oscura!!!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("¡¡¡Has sido expulsado de la Legión Oscura!!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje093 ' "¡¡¡Te has retirado de la Legión Oscura!!!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("¡¡¡Te has retirado de la Legión Oscura!!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje094 ' "Hoy es la votación para elegir un nuevo líder para el clan."
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("Hoy es la votación para elegir un nuevo líder para el clan.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje095 ' "La elección durará 24 horas, se puede votar a cualquier miembro del clan."
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("La elección durará 24 horas, se puede votar a cualquier miembro del clan.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje096 ' "Para votar escribe /VOTO NICKNAME."
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("Para votar escribe /VOTO NICKNAME.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje097 ' "Sólo se computará un voto por miembro. Tu voto no puede ser cambiado."
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("Sólo se computará un voto por miembro. Tu voto no puede ser cambiado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje098 ' "Error, el clan no existe."
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("Error, el clan no existe.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje099 ' "No perteneces a ningún clan."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No perteneces a ningún clan.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje100 ' "No eres el líder de tu clan."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No eres el líder de tu clan.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje101 ' "El personaje no es ni aspirante ni miembro del clan."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El personaje no es ni aspirante ni miembro del clan.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje102 ' "No está permitido arrojar objetos al suelo en zonas seguras.", FontTypeNames.FONTTYPE_CITIZEN
      With FontTypes(FontTypeNames.FONTTYPE_CITIZEN)
         Call ShowConsoleMsg("No está permitido arrojar objetos al suelo en zonas seguras.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje103 ' "No tienes espacio para más hechizos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes espacio para más hechizos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje104 ' "Ya sabes el hechizo."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("Ya sabes el hechizo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje105 ' "No puedes lanzar hechizos estando muerto."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes lanzar hechizos estando muerto.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje106 ' "No posees un báculo lo suficientemente poderoso para poder lanzar el conjuro."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No posees un báculo lo suficientemente poderoso para poder lanzar el conjuro.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje107 ' "No puedes lanzar este conjuro sin la ayuda de un báculo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes lanzar este conjuro sin la ayuda de un báculo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje108 ' "No tienes suficientes puntos de magia para lanzar este hechizo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes suficientes puntos de magia para lanzar este hechizo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje109 ' "Estás muy cansado para lanzar este hechizo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Estás muy cansado para lanzar este hechizo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje110 ' "Estás muy cansada para lanzar este hechizo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Estás muy cansada para lanzar este hechizo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje111 ' "Debes poseer toda tu maná para poder lanzar este hechizo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Debes poseer toda tu maná para poder lanzar este hechizo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje112 ' "Debes poseer alguna mascota para poder lanzar este hechizo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Debes poseer alguna mascota para poder lanzar este hechizo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje113 ' "No tienes suficiente maná."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("No tienes suficiente maná.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje114 ' "No puedes invocar criaturas en zona segura."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes invocar criaturas en zona segura.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje115 ' "No puedes lanzar hechizos si estás en consulta."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes lanzar hechizos si estás en consulta.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje116 ' "Estás demasiado lejos para lanzar este hechizo."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("Estás demasiado lejos para lanzar este hechizo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje117 ' "Este hechizo actúa sólo sobre usuarios."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Este hechizo actúa sólo sobre usuarios.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje118 ' "Este hechizo sólo afecta a los npcs."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Este hechizo sólo afecta a los npcs.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje119 ' "Target inválido."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Target inválido.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje120 ' "¡El usuario está muerto!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡El usuario está muerto!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje121 ' "¡El hechizo no tiene efecto!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡El hechizo no tiene efecto!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje122 ' "¡No puedes hacerte invisible mientras te encuentras saliendo!"
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("¡No puedes hacerte invisible mientras te encuentras saliendo!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje123 ' "¡La invisibilidad no funciona aquí!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡La invisibilidad no funciona aquí!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje124 ' "Ya te encuentras mimetizado. El hechizo no ha tenido efecto."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Ya te encuentras mimetizado. El hechizo no ha tenido efecto.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje125 ' "No puedes atacarte a vos mismo."
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("No puedes atacarte a vos mismo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje126 ' " ¡El hechizo no tiene efecto!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg(" ¡El hechizo no tiene efecto!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje127 ' "¡El espíritu no tiene intenciones de regresar al mundo de los vivos!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡El espíritu no tiene intenciones de regresar al mundo de los vivos!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje128 ' "¡Revivir no está permitido aquí! Retirate de la Zona si deseas utilizar el Hechizo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Revivir no está permitido aquí! Retirate de la Zona si deseas utilizar el Hechizo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje129 ' "No puedes resucitar si no tienes tu barra de energía llena."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes resucitar si no tienes tu barra de energía llena.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje130 ' "Necesitas un báculo mejor para lanzar este hechizo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Necesitas un báculo mejor para lanzar este hechizo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje131 ' "Necesitas un instrumento mágico para devolver la vida."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Necesitas un instrumento mágico para devolver la vida.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje132 ' "¡Los Dioses te sonríen, has ganado 500 puntos de nobleza!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Los Dioses te sonríen, has ganado 500 puntos de nobleza!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje133 ' "El esfuerzo de resucitar fue demasiado grande."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El esfuerzo de resucitar fue demasiado grande.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje134 ' "El esfuerzo de resucitar te ha debilitado."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El esfuerzo de resucitar te ha debilitado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje135 ' "Tu viaje ha sido cancelado."
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Tu viaje ha sido cancelado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje136 ' "El NPC es inmune a este hechizo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El NPC es inmune a este hechizo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje137 ' "Sólo puedes remover la parálisis de los Guardias si perteneces a su facción."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Sólo puedes remover la parálisis de los Guardias si perteneces a su facción.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje138 ' "Solo puedes remover la parálisis de los NPCs que te consideren su amo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Solo puedes remover la parálisis de los NPCs que te consideren su amo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje139 ' "Solo puedes remover la parálisis de los Guardias si perteneces a su facción."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Solo puedes remover la parálisis de los Guardias si perteneces a su facción.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje140 ' "Este NPC no está paralizado"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Este NPC no está paralizado", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje141 ' "El NPC es inmune al hechizo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El NPC es inmune al hechizo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje142 ' "Sólo los druidas pueden mimetizarse con criaturas."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Sólo los druidas pueden mimetizarse con criaturas.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje143 ' "No puedes lanzar este hechizo a un muerto."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes lanzar este hechizo a un muerto.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje144 ' "No puedes ayudar usuarios mientras estas en consulta."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes ayudar usuarios mientras estas en consulta.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje145 ' "Los miembros del ejército real no pueden ayudar a los criminales."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Los miembros del ejército real no pueden ayudar a los criminales.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje146 ' "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje147 ' "Los miembros de la legión oscura no pueden ayudar a los ciudadanos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Los miembros de la legión oscura no pueden ayudar a los ciudadanos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje148 ' "Los miembros del ejército real no pueden ayudar a ciudadanos en estado atacable."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Los miembros del ejército real no pueden ayudar a ciudadanos en estado atacable.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje149 ' "Para ayudar ciudadanos en estado atacable debes sacarte el seguro, pero te puedes volver criminal."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Para ayudar ciudadanos en estado atacable debes sacarte el seguro, pero te puedes volver criminal.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje150 ' "No puedes mover el hechizo en esa dirección."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes mover el hechizo en esa dirección.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje151 ' "¡Has matado a la criatura!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("¡Has matado a la criatura!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje152 ' "¡¡La criatura te ha envenenado!!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("¡¡La criatura te ha envenenado!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje153 ' "¡Has subido de nivel!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Has subido de nivel!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje154 ' "¡Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearás! Por esta razón, hasta tanto no te enlistes en la facción bajo la cual tu clan está alineado, estarás excluído del mismo."
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("¡Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearás! Por esta razón, hasta tanto no te enlistes en la facción bajo la cual tu clan está alineado, estarás excluído del mismo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje155 ' "Debes abandonar el Dungeon Newbie."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Debes abandonar el Dungeon Newbie.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje156 ' "(CUERPO) Mín Def/Máx Def: 0"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("(CUERPO) Mín Def/Máx Def: 0", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje157 ' "(CABEZA) Mín Def/Máx Def: 0"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("(CABEZA) Mín Def/Máx Def: 0", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje158 ' "Status: Líder"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Status: Líder", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje159 ' "Fue ejército real"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Fue ejército real", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje160 ' "Fue legión oscura"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Fue legión oscura", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje161 ' "Para poder entrenar un skill debes asignar los 10 skills iniciales."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Para poder entrenar un skill debes asignar los 10 skills iniciales.", .Red, .Green, .Blue, .bold, .italic)
      End With
      ' GSZAO Abrimos la ventana para asignar skills automaticamente!
      Call frmMain.RequestAsignarSkills
   Case Is = eMensajes.Mensaje162 ' "¡Has ganado 50 puntos de experiencia!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("¡Has ganado 50 puntos de experiencia!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje163 ' "Tus mascotas no pueden transitar este mapa."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Tus mascotas no pueden transitar este mapa.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje164 ' "Pierdes el control de tus mascotas invocadas."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Pierdes el control de tus mascotas invocadas.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje165 ' "No se permiten mascotas en zona segura. Éstas te esperarán afuera."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No se permiten mascotas en zona segura. Éstas te esperarán afuera.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje166 ' "Tu mascota no pueden transitar este sector del mapa, intenta invocarla en otra parte."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Tu mascota no pueden transitar este sector del mapa, intenta invocarla en otra parte.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje167 ' "¡Has recuperado tu apariencia normal!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Has recuperado tu apariencia normal!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje168 ' "/salir cancelado."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("/salir cancelado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje169 ' "Personaje Inexistente"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Personaje Inexistente", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje170 ' "Debes estar muerto para poder utilizar este comando."
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Debes estar muerto para poder utilizar este comando.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje171 ' "No puedes robar npcs en zonas seguras."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes robar npcs en zonas seguras.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje172 ' "No puedes atacar otra criatura con dueño hasta que haya terminado tu castigo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes atacar otra criatura con dueño hasta que haya terminado tu castigo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje173 ' "El rey pretoriano te ha vuelto ciego "
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("El rey pretoriano te ha vuelto ciego ", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje174 ' "A la distancia escuchas las siguientes palabras: ¡Cobarde, no eres digno de luchar conmigo si escapas! "
      With FontTypes(FontTypeNames.FONTTYPE_VENENO)
         Call ShowConsoleMsg("A la distancia escuchas las siguientes palabras: ¡Cobarde, no eres digno de luchar conmigo si escapas! ", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje175 ' "El rey pretoriano te ha vuelto estúpido."
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("El rey pretoriano te ha vuelto estúpido.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje176 ' "¡Has sido detectado!"
      With FontTypes(FontTypeNames.FONTTYPE_VENENO)
         Call ShowConsoleMsg("¡Has sido detectado!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje177 ' "Comienzas a hacerte visible."
      With FontTypes(FontTypeNames.FONTTYPE_VENENO)
         Call ShowConsoleMsg("Comienzas a hacerte visible.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje178 ' "Ya te encuentras en tu hogar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Ya te encuentras en tu hogar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje179 ' "No puedes usar este comando aquí."
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("No puedes usar este comando aquí.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje180 ' "Debes estar muerto para utilizar este comando."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Debes estar muerto para utilizar este comando.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje181 ' "¡Has vuelto a ser visible!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Has vuelto a ser visible!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje182 ' "¡¡Estás muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. "
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡Estás muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. ", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje183 ' "Usuario inexistente."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Usuario inexistente.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje184 ' "No puedes susurrarle a los Dioses y Admins."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes susurrarle a los Dioses y Admins.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje185 ' "No puedes susurrarle a los GMs."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes susurrarle a los GMs.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje186 ' "Estás muy lejos del usuario."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Estás muy lejos del usuario.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje187 ' "Dejas de meditar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Dejas de meditar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje188 ' "Has dejado de descansar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has dejado de descansar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje189 ' "No puedes moverte porque estás paralizado."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes moverte porque estás paralizado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje190 ' "No puedes usar así este arma."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes usar así este arma.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje191 ' "No puedes tomar ningún objeto."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes tomar ningún objeto.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje192 ' "Has dejado de comerciar."
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("Has dejado de comerciar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje193 ' "Has rechazado la oferta del otro usuario."
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("Has rechazado la oferta del otro usuario.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje194 ' "No puedes ocultarte si estás en consulta."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes ocultarte si estás en consulta.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje195 ' "No puedes ocultarte si estás navegando."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes ocultarte si estás navegando.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje196 ' "Ya estás oculto."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Ya estás oculto.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje197 ' "No tienes municiones."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes municiones.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje198 ' "Estás muy cansado para luchar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Estás muy cansado para luchar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje199 ' "Estás muy cansada para luchar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Estás muy cansada para luchar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje200 ' "Estás demasiado lejos para atacar."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("Estás demasiado lejos para atacar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje201 ' "¡No puedes atacarte a vos mismo!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡No puedes atacarte a vos mismo!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje202 ' "Una fuerza oscura te impide canalizar tu energía."
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Una fuerza oscura te impide canalizar tu energía.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje203 ' "¡Primero selecciona el hechizo que quieres lanzar!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Primero selecciona el hechizo que quieres lanzar!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje204 ' "No puedes pescar desde donde te encuentras."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes pescar desde donde te encuentras.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje205 ' "Estás demasiado lejos para pescar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Estás demasiado lejos para pescar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje206 ' "No hay agua donde pescar. Busca un lago, río o mar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay agua donde pescar. Busca un lago, río o mar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje207 ' "No puedes robar aquí."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("No puedes robar aquí.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje208 ' "¡No hay a quien robarle!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡No hay a quien robarle!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje209 ' "¡No puedes robar en zonas seguras!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡No puedes robar en zonas seguras!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje210 ' "Deberías equiparte el hacha."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Deberías equiparte el hacha.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje211 ' "No puedes talar desde allí."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes talar desde allí.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje212 ' "El hacha utilizado no es suficientemente poderosa."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El hacha utilizado no es suficientemente poderosa.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje213 ' "No hay ningún árbol ahí."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay ningún árbol ahí.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje214 ' "Ahí no hay ningún yacimiento."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Ahí no hay ningún yacimiento.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje215 ' "No puedes domar una criatura que está luchando con un jugador."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes domar una criatura que está luchando con un jugador.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje216 ' "No puedes domar a esa criatura."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes domar a esa criatura.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje217 ' "¡No hay ninguna criatura allí!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡No hay ninguna criatura allí!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje218 ' "No tienes más minerales."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes más minerales.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje219 ' "Ahí no hay ninguna fragua."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Ahí no hay ninguna fragua.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje220 ' "Ahí no hay ningún yunque."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Ahí no hay ningún yunque.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje221 ' "¡Primero selecciona el hechizo!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Primero selecciona el hechizo!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje222 ' "No estás comerciando."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No estás comerciando.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje223 ' "Propuesta de paz enviada."
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("Propuesta de paz enviada.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje224 ' "Propuesta de alianza enviada."
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("Propuesta de alianza enviada.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje225 ' "El personaje no ha mandado solicitud, o no estás habilitado para verla."
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("El personaje no ha mandado solicitud, o no estás habilitado para verla.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje226 ' "No puedes expulsar ese personaje del clan."
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("No puedes expulsar ese personaje del clan.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje227 ' "No puedes salir estando paralizado."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("No puedes salir estando paralizado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje228 ' "Comercio cancelado."
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("Comercio cancelado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje229 ' "Dejas el clan."
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("Dejas el clan.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje230 ' "Tú no puedes salir de este clan."
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("Tú no puedes salir de este clan.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje231 ' "Primero tienes que seleccionar un NPC, haz click izquierdo sobre él."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Primero tienes que seleccionar un NPC, haz click izquierdo sobre él.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje232 ' "¡¡Estás muerto!! Solo puedes usar ítems cuando estás vivo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡Estás muerto!! Solo puedes usar ítems cuando estás vivo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje233 ' "Te acomodás junto a la fogata y comienzas a descansar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Te acomodás junto a la fogata y comienzas a descansar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje234 ' "Te levantas."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Te levantas.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje235 ' "No hay ninguna fogata junto a la cual descansar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay ninguna fogata junto a la cual descansar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje236 ' "¡¡Estás muerto!! Sólo puedes meditar cuando estás vivo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡Estás muerto!! Sólo puedes meditar cuando estás vivo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje237 ' "Sólo las clases mágicas conocen el arte de la meditación."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Sólo las clases mágicas conocen el arte de la meditación.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje238 ' "Maná restaurado."
      With FontTypes(FontTypeNames.FONTTYPE_VENENO)
         Call ShowConsoleMsg("Maná restaurado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje239 ' "El sacerdote no puede resucitarte debido a que estás demasiado lejos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El sacerdote no puede resucitarte debido a que estás demasiado lejos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje240 ' "¡¡Has sido resucitado!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡Has sido resucitado!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje241 ' "Primero tienes que seleccionar un usuario, haz click izquierdo sobre él."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Primero tienes que seleccionar un usuario, haz click izquierdo sobre él.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje242 ' "No puedes iniciar el modo consulta con otro administrador."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes iniciar el modo consulta con otro administrador.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje243 ' "Has terminado el modo consulta."
      With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
         Call ShowConsoleMsg("Has terminado el modo consulta.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje244 ' "Has iniciado el modo consulta."
      With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
         Call ShowConsoleMsg("Has iniciado el modo consulta.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje245 ' "¡¡Has sido curado!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡Has sido curado!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje246 ' "Ya estás comerciando."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Ya estás comerciando.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje247 ' "¡¡No puedes comerciar con los muertos!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡No puedes comerciar con los muertos!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje248 ' "¡¡No puedes comerciar con vos mismo!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡¡No puedes comerciar con vos mismo!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje249 ' "Estás demasiado lejos del usuario."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Estás demasiado lejos del usuario.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje250 ' "No puedes comerciar con el usuario en este momento."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes comerciar con el usuario en este momento.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje251 ' "Primero haz click izquierdo sobre el personaje."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Primero haz click izquierdo sobre el personaje.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje252 ' "Debes acercarte más."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Debes acercarte más.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje253 ' "No puedes compartir npcs con administradores!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes compartir npcs con administradores!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje254 ' "Solo puedes compartir npcs con miembros de tu misma facción!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Solo puedes compartir npcs con miembros de tu misma facción!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje255 ' "No puedes compartir npcs con criminales!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes compartir npcs con criminales!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje256 ' "No pertences a ningún clan."
      With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
         Call ShowConsoleMsg("No pertences a ningún clan.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje257 ' "Su solicitud ha sido enviada."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Su solicitud ha sido enviada.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje258 ' "El mensaje ha sido entregado, ahora sólo debes esperar que se desocupe algún GM."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El mensaje ha sido entregado, ahora sólo debes esperar que se desocupe algún GM.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje259 ' "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje260 ' "No puedes cambiar la descripción estando muerto."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes cambiar la descripción estando muerto.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje261 ' "La descripción tiene caracteres inválidos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("La descripción tiene caracteres inválidos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje262 ' "La descripción ha cambiado."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("La descripción ha cambiado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje263 ' "Voto contabilizado."
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("Voto contabilizado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje264 ' "No puedes ver las penas de los administradores."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes ver las penas de los administradores.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje265 ' "Sin prontuario.."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Sin prontuario..", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje266 ' "Debes especificar una contraseña nueva, inténtalo de nuevo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Debes especificar una contraseña nueva, inténtalo de nuevo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje267 ' "La contraseña actual proporcionada no es correcta. La contraseña no ha sido cambiada, inténtalo de nuevo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("La contraseña actual proporcionada no es correcta. La contraseña no ha sido cambiada, inténtalo de nuevo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje268 ' "La contraseña fue cambiada con éxito."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("La contraseña fue cambiada con éxito.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje269 ' "¡No perteneces a ninguna facción!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("¡No perteneces a ninguna facción!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje270 ' "Denuncia enviada, espere.."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Denuncia enviada, espere..", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje271 ' "¡Ya has fundado un clan, no puedes fundar otro!"
      With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
         Call ShowConsoleMsg("¡Ya has fundado un clan, no puedes fundar otro!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje272 ' "Alineación inválida."
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("Alineación inválida.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje273 ' "No puedes incorporar a tu party a personajes de mayor jerarquía."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes incorporar a tu party a personajes de mayor jerarquía.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje274 ' "No hay reales conectados."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay reales conectados.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje275 ' "No hay Caos conectados."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay Caos conectados.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje276 ' "Usuario offline."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Usuario offline.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje277 ' "Todos los lugares están ocupados."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Todos los lugares están ocupados.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje278 ' "Comentario salvado..."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Comentario salvado...", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje279 ' "Npcs Hostiles en mapa: "
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("Npcs Hostiles en mapa: ", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje280 ' "No hay NPCS Hostiles."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay NPCS Hostiles.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje281 ' "Otros Npcs en mapa: "
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("Otros Npcs en mapa: ", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje282 ' "No hay más NPCS."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay más NPCS.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje283 ' "Usuario silenciado."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Usuario silenciado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje284 ' "Usuario des silenciado."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Usuario des silenciado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje285 ' "No perteneces a ningún grupo!"
      With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
         Call ShowConsoleMsg("No perteneces a ningún grupo!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje286 ' "No hay usuarios trabajando."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay usuarios trabajando.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje287 ' "No hay usuarios ocultandose."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay usuarios ocultandose.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje288 ' "Utilice /carcel nick@motivo@tiempo"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Utilice /carcel nick@motivo@tiempo", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje289 ' "No puedes encarcelar a administradores."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes encarcelar a administradores.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje290 ' "No puedés encarcelar por más de 60 minutos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedés encarcelar por más de 60 minutos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje291 ' "Los consejeros no pueden usar este comando en el mapa pretoriano."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Los consejeros no pueden usar este comando en el mapa pretoriano.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje292 ' "Antes debes hacer click sobre el NPC."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Antes debes hacer click sobre el NPC.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje293 ' "Utilice /advertencia nick@motivo"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Utilice /advertencia nick@motivo", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje294 ' "No puedes advertir a administradores."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes advertir a administradores.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje295 ' "Estás intentando editar un usuario inexistente."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Estás intentando editar un usuario inexistente.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje296 ' "Clase desconocida. Intente nuevamente."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Clase desconocida. Intente nuevamente.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje297 ' "Skill Inexistente!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Skill Inexistente!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje298 ' "Genero desconocido. Intente nuevamente."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Genero desconocido. Intente nuevamente.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje299 ' "Raza desconocida. Intente nuevamente."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Raza desconocida. Intente nuevamente.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje300 ' "Comando no permitido."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Comando no permitido.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje301 ' "Usuario offline, buscando en charfile."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Usuario offline, buscando en charfile.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje302 ' "Usuario offline. Leyendo charfile... "
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Usuario offline. Leyendo charfile... ", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje303 ' "Usuario offline. Leyendo del charfile..."
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("Usuario offline. Leyendo del charfile...", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje304 ' "No hay GMs Online."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay GMs Online.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje305 ' "Sólo se permite perdonar newbies."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Sólo se permite perdonar newbies.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje306 ' "No puedes echar a alguien con jerarquía mayor a la tuya."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes echar a alguien con jerarquía mayor a la tuya.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje307 ' "¿¿Estás loco?? ¿¿Cómo vas a piñatear un gm?? :@"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¿¿Estás loco?? ¿¿Cómo vas a piñatear un gm?? :@", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje308 ' "No está online."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No está online.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje309 ' "Charfile inexistente (no use +)."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Charfile inexistente (no use +).", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje310 ' "El jugador no está online."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El jugador no está online.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje311 ' "No puedes invocar a dioses y admins."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes invocar a dioses y admins.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje312 ' "No hay ningún personaje con ese nick."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay ningún personaje con ese nick.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje313 ' "Hay un objeto en el piso en ese lugar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Hay un objeto en el piso en ese lugar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje314 ' "No puedes crear un teleport que apunte a la entrada de otro."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes crear un teleport que apunte a la entrada de otro.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje315 ' "Haz click sobre un personaje antes."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Haz click sobre un personaje antes.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje316 ' "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Debes seleccionar el NPC por el que quieres hablar antes de usar este comando.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje317 ' "Usuario offline"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Usuario offline", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje318 ' "Usuario offline, echando de los consejos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Usuario offline, echando de los consejos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje319 ' "Has sido echado del consejo de Banderbill."
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("Has sido echado del consejo de Banderbill.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje320 ' "Has sido echado del Concilio de las Sombras."
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("Has sido echado del Concilio de las Sombras.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje321 ' "El personaje no está online."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El personaje no está online.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje322 ' "¡¡ATENCIÓN: FUERON CREADOS ***100*** ÍTEMS, TIRE Y /DEST LOS QUE NO NECESITE!!"
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("¡¡ATENCIÓN: FUERON CREADOS ***100*** ÍTEMS, TIRE Y /DEST LOS QUE NO NECESITE!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje323 ' "No puede destruir teleports así. Utilice /DT."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puede destruir teleports así. Utilice /DT.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje324 ' "Utilice /borrarpena Nick@NumeroDePena@NuevaPena"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Utilice /borrarpena Nick@NumeroDePena@NuevaPena", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje325 ' "Pena modificada."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Pena modificada.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje326 ' "No hay ningún objeto en slot seleccionado."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No hay ningún objeto en slot seleccionado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje327 ' "Slot Inválido."
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("Slot Inválido.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje328 ' "Npcs.dat recargado."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Npcs.dat recargado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje329 ' "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje330 ' "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje331 ' "Igualmente, el único útil es 'NIEVE' ya que al ingresarlo, la gente muere de frío en el mapa."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Igualmente, el único útil es 'NIEVE' ya que al ingresarlo, la gente muere de frío en el mapa.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje332 ' "Igualmente, el único útil es 'DUNGEON' ya que al ingresarlo, NO se sentirá el efecto de la lluvia en este mapa."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Igualmente, el único útil es 'DUNGEON' ya que al ingresarlo, NO se sentirá el efecto de la lluvia en este mapa.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje333 ' "Mapa Guardado."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Mapa Guardado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje334 ' "Usar: /ANAME origen@destino"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Usar: /ANAME origen@destino", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje335 ' "El Pj está online, debe salir para hacer el cambio."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("El Pj está online, debe salir para hacer el cambio.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje336 ' "Transferencia exitosa."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Transferencia exitosa.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje337 ' "El nick solicitado ya existe."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El nick solicitado ya existe.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje338 ' "usar /AEMAIL <pj>-<nuevomail>"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("usar /AEMAIL <pj>-<nuevomail>", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje339 ' "usar /APASS <pjsinpass>@<pjconpass>"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("usar /APASS <pjsinpass>@<pjconpass>", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje340 ' "Servidor habilitado para todos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Servidor habilitado para todos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje341 ' "Servidor restringido a administradores."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Servidor restringido a administradores.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje342 ' "No pertenece a ningún clan o es fundador."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No pertenece a ningún clan o es fundador.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje343 ' "Expulsado."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Expulsado.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje344 ' "Se ha cambiado el MOTD con éxito."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Se ha cambiado el MOTD con éxito.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje345 ' "¡No puedes modificar esa información desde aquí!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡No puedes modificar esa información desde aquí!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje346 ' "No existe la llave y/o clave"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No existe la llave y/o clave", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje347 ' "Debes matar al resto del ejército antes de atacar al rey!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Debes matar al resto del ejército antes de atacar al rey!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje348 ' "No puedes atacar mascotas en zona segura."
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("No puedes atacar mascotas en zona segura.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje349 ' "No puedes atacar a este NPC."
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("No puedes atacar a este NPC.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje350 ' "Estás muy lejos para disparar."
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Estás muy lejos para disparar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje351 ' "No puedes atacar a un espíritu."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes atacar a un espíritu.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje352 ' "No puedes atacar usuarios mientras estas en consulta."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes atacar usuarios mientras estas en consulta.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje353 ' "No puedes atacar usuarios mientras estan en consulta."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes atacar usuarios mientras estan en consulta.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje354 ' "El ser es demasiado poderoso."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("El ser es demasiado poderoso.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje355 ' "Los soldados del ejército real tienen prohibido atacar ciudadanos."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("Los soldados del ejército real tienen prohibido atacar ciudadanos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje356 ' "Los miembros de la legión oscura tienen prohibido atacarse entre sí."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("Los miembros de la legión oscura tienen prohibido atacarse entre sí.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje357 ' "No puedes atacar ciudadanos, para hacerlo debes desactivar el seguro."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("No puedes atacar ciudadanos, para hacerlo debes desactivar el seguro.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje358 ' "¡Huye de la ciudad! Estás siendo atacado y no podrás defenderte."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("¡Huye de la ciudad! Estás siendo atacado y no podrás defenderte.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje359 ' "Esta es una zona segura, aquí no puedes atacar a otros usuarios."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("Esta es una zona segura, aquí no puedes atacar a otros usuarios.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje360 ' "No puedes pelear aquí."
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("No puedes pelear aquí.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje361 ' "No puedes atacar npcs mientras estas en consulta."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes atacar npcs mientras estas en consulta.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje362 ' "No puedes atacar esta criatura."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes atacar esta criatura.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje363 ' "No puedes atacar Guardias del Caos siendo de la legión oscura."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes atacar Guardias del Caos siendo de la legión oscura.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje364 ' "No puedes atacar Guardias Reales siendo del ejército real."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes atacar Guardias Reales siendo del ejército real.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje365 ' "Para poder atacar Guardias Reales debes quitarte el seguro."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Para poder atacar Guardias Reales debes quitarte el seguro.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje366 ' "¡Atacaste un Guardia Real! Eres un criminal."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Atacaste un Guardia Real! Eres un criminal.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje367 ' "Los miembros del ejército real no pueden atacar npcs no hostiles."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Los miembros del ejército real no pueden atacar npcs no hostiles.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje368 ' "Para atacar a este NPC debes quitarte el seguro."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Para atacar a este NPC debes quitarte el seguro.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje369 ' "Atacaste un NPC no-hostil. Continúa haciéndolo y te podrás convertir en criminal."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Atacaste un NPC no-hostil. Continúa haciéndolo y te podrás convertir en criminal.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje370 ' "Los miembros del ejército real no pueden atacar mascotas de ciudadanos."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Los miembros del ejército real no pueden atacar mascotas de ciudadanos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje371 ' "Para atacar mascotas de ciudadanos debes quitarte el seguro."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Para atacar mascotas de ciudadanos debes quitarte el seguro.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje372 ' "Has atacado la Mascota de un ciudadano. Eres un criminal."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has atacado la Mascota de un ciudadano. Eres un criminal.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje373 ' "Los miembros de la legión oscura no pueden atacar mascotas de otros legionarios. "
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Los miembros de la legión oscura no pueden atacar mascotas de otros legionarios. ", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje374 ' "Los miembros del Ejército Real no pueden paralizar criaturas ya paralizadas pertenecientes a otros miembros del Ejército Real"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Los miembros del Ejército Real no pueden paralizar criaturas ya paralizadas pertenecientes a otros miembros del Ejército Real", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje375 ' "Para paralizar criaturas ya paralizadas pertenecientes a ciudadanos debes quitarte el seguro."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Para paralizar criaturas ya paralizadas pertenecientes a ciudadanos debes quitarte el seguro.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje376 ' "Has paralizado la criatura de un ciudadano, ahora eres atacable por él."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has paralizado la criatura de un ciudadano, ahora eres atacable por él.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje377 ' "Los miembros de la legión oscura no pueden paralizar criaturas ya paralizadas por otros legionarios."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Los miembros de la legión oscura no pueden paralizar criaturas ya paralizadas por otros legionarios.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje378 ' "Los miembros del Ejército Real no pueden atacar criaturas pertenecientes a otros miembros del Ejército Real"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Los miembros del Ejército Real no pueden atacar criaturas pertenecientes a otros miembros del Ejército Real", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje379 ' "Para atacar criaturas ya pertenecientes a ciudadanos debes quitarte el seguro."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Para atacar criaturas ya pertenecientes a ciudadanos debes quitarte el seguro.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje380 ' "Has atacado a la criatura de un ciudadano, ahora eres atacable por él."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has atacado a la criatura de un ciudadano, ahora eres atacable por él.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje381 ' "Para atacar criaturas pertenecientes a ciudadanos debes quitarte el seguro."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Para atacar criaturas pertenecientes a ciudadanos debes quitarte el seguro.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje382 ' "Los miembros de la Legión Oscura no pueden atacar criaturas de otros legionarios. "
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Los miembros de la Legión Oscura no pueden atacar criaturas de otros legionarios. ", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje383 ' "Debes matar al resto del ejército antes de atacar al rey."
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Debes matar al resto del ejército antes de atacar al rey.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje384 ' "Comercio cancelado por el otro usuario"
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("Comercio cancelado por el otro usuario", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje385 ' "Servidor> Por favor espera algunos segundos, el WorldSave está ejecutándose."
      With FontTypes(FontTypeNames.FONTTYPE_SERVER)
         Call ShowConsoleMsg("Servidor> Por favor espera algunos segundos, el WorldSave está ejecutándose.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje386 ' "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde."
      With FontTypes(FontTypeNames.FONTTYPE_SERVER)
         Call ShowConsoleMsg("Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje387 ' "Tu estado no te permite entrar al clan."
      With FontTypes(FontTypeNames.FONTTYPE_GUILD)
         Call ShowConsoleMsg("Tu estado no te permite entrar al clan.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje388 ' "¡Te has escondido entre las sombras!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Te has escondido entre las sombras!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje389 ' "¡No has logrado esconderte!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡No has logrado esconderte!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje390 ' "No tienes suficientes conocimientos para usar este barco."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes suficientes conocimientos para usar este barco.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje391 ' "No tienes conocimientos de minería suficientes para trabajar este mineral."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes conocimientos de minería suficientes para trabajar este mineral.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje392 ' "No tienes los conocimientos suficientes en herrería para fundir este objeto."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes los conocimientos suficientes en herrería para fundir este objeto.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje393 ' "No tienes suficiente madera."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes suficiente madera.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje394 ' "No tienes suficiente madera élfica."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes suficiente madera élfica.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje395 ' "No tienes suficientes lingotes de hierro."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes suficientes lingotes de hierro.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje396 ' "No tienes suficientes lingotes de plata."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes suficientes lingotes de plata.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje397 ' "No tienes suficientes lingotes de oro."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes suficientes lingotes de oro.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje398 ' "No tienes suficientes materiales."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes suficientes materiales.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje399 ' "No tienes suficiente energía."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes suficiente energía.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje400 ' "Debes tener equipado el serrucho para trabajar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Debes tener equipado el serrucho para trabajar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje401 ' "No tienes suficientes minerales para hacer un lingote."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes suficientes minerales para hacer un lingote.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje402 ' "Debes equiparte el martillo de herrero."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Debes equiparte el martillo de herrero.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje403 ' "No tienes suficientes skills."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No tienes suficientes skills.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje404 ' "Has mejorado el arma!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has mejorado el arma!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje405 ' "Has mejorado el escudo!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has mejorado el escudo!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje406 ' "Has mejorado el casco!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has mejorado el casco!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje407 ' "Has mejorado la armadura!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has mejorado la armadura!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje408 ' "Debes equiparte el serrucho."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Debes equiparte el serrucho.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje409 ' "Has mejorado la flecha!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has mejorado la flecha!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje410 ' "Has mejorado el barco!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has mejorado el barco!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje411 ' "Ya domaste a esa criatura."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Ya domaste a esa criatura.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje412 ' "La criatura ya tiene amo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("La criatura ya tiene amo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje413 ' "No puedes domar más de dos criaturas del mismo tipo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes domar más de dos criaturas del mismo tipo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje414 ' "La criatura te ha aceptado como su amo."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("La criatura te ha aceptado como su amo.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje415 ' "No has logrado domar la criatura."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No has logrado domar la criatura.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje416 ' "No puedes controlar más criaturas."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes controlar más criaturas.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje417 ' "Necesitas clickear sobre leña para hacer ramitas."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Necesitas clickear sobre leña para hacer ramitas.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje418 ' "Estás demasiado lejos para prender la fogata."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Estás demasiado lejos para prender la fogata.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje419 ' "No puedes hacer fogatas estando muerto."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes hacer fogatas estando muerto.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje420 ' "Necesitas por lo menos tres troncos para hacer una fogata."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Necesitas por lo menos tres troncos para hacer una fogata.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje421 ' "No has podido hacer la fogata."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No has podido hacer la fogata.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje422 ' "¡Has pescado un lindo pez!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Has pescado un lindo pez!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje423 ' "¡No has pescado nada!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡No has pescado nada!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje424 ' "¡Has pescado algunos peces!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Has pescado algunos peces!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje425 ' "Debes quitarte el seguro para robarle a un ciudadano."
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Debes quitarte el seguro para robarle a un ciudadano.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje426 ' "Los miembros del ejército real no tienen permitido robarle a ciudadanos."
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Los miembros del ejército real no tienen permitido robarle a ciudadanos.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje427 ' "No puedes robar a otros miembros de la legión oscura."
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("No puedes robar a otros miembros de la legión oscura.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje428 ' "Estás muy cansado para robar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Estás muy cansado para robar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje429 ' "Estás muy cansada para robar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Estás muy cansada para robar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje430 ' "¡No has logrado robar nada!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡No has logrado robar nada!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje431 ' "No has logrado robar ningún objeto."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No has logrado robar ningún objeto.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje432 ' "¡No has logrado apuñalar a tu enemigo!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("¡No has logrado apuñalar a tu enemigo!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje433 ' "¡Has conseguido algo de leña!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Has conseguido algo de leña!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje434 ' "¡No has obtenido leña!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡No has obtenido leña!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje435 ' "¡Has extraido algunos minerales!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡Has extraido algunos minerales!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje436 ' "¡No has conseguido nada!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡No has conseguido nada!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje437 ' "Has terminado de meditar."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has terminado de meditar.", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje438 ' "Has logrado desequipar el escudo de tu oponente!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Has logrado desequipar el escudo de tu oponente!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje439 ' "¡Tu oponente te ha desequipado el escudo!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("¡Tu oponente te ha desequipado el escudo!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje440 ' "Has logrado desarmar a tu oponente!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Has logrado desarmar a tu oponente!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje441 ' "¡Tu oponente te ha desarmado!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("¡Tu oponente te ha desarmado!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje442 ' "Has logrado desequipar el casco de tu oponente!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Has logrado desequipar el casco de tu oponente!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje443 ' "¡Tu oponente te ha desequipado el casco!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("¡Tu oponente te ha desequipado el casco!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje444 ' "Tu oponente no tiene equipado items!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Tu oponente no tiene equipado items!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje445 ' "No has logrado desequipar ningún item a tu oponente!"
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("No has logrado desequipar ningún item a tu oponente!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje446 ' "Tu golpe ha dejado inmóvil a tu oponente"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Tu golpe ha dejado inmóvil a tu oponente", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje447 ' "¡El golpe te ha dejado inmóvil!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("¡El golpe te ha dejado inmóvil!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje448 ' "Operación realizada con exito!!"
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Operación realizada con exito!!", .Red, .Green, .Blue, .bold, .italic)
      End With
   Case Is = eMensajes.Mensaje449 ' "El usuario no se encuentra en el listado solicitado."
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El usuario no se encuentra en el listado solicitado.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje450 ' "Para recorrer los mares debes ser nivel 20 y además tu skill en pesca debe ser 100." *-*  FontTypeNames.FONTTYPE_INFO
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Para recorrer los mares debes ser nivel 20 y además tu skill en pesca debe ser 100.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje451 ' "Para recorrer los mares debes ser nivel 20 o superior." *-*  FontTypeNames.FONTTYPE_INFO
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Para recorrer los mares debes ser nivel 20 o superior.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje452 ' "No puedes comerciar en este momento" *-*  FontTypeNames.FONTTYPE_TALK
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("No puedes comerciar en este momento.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje453 ' "¡Los miembros del staff no pueden crear partys!" *-*  FontTypeNames.FONTTYPE_TALK
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("¡Los miembros del staff no pueden crear partys!", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje454 ' "¡Los miembros del staff no pueden unirse a partys!" *-*  FontTypeNames.FONTTYPE_TALK
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("¡Los miembros del staff no pueden unirse a partys!", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje455 ' "Invocar no está permitido aquí! Retirate de la Zona si deseas utilizar el Hechizo." *-*  FontTypeNames.FONTTYPE_TALK
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("Invocar no está permitido aquí! Retirate de la Zona si deseas utilizar el Hechizo.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje456 ' "¡¡¡Tu estado no te permite permanecer en el mapa!!!" *-*  FontTypeNames.FONTTYPE_INFOBOLD
      With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
         Call ShowConsoleMsg("¡¡¡Tu estado no te permite permanecer en el mapa!!!", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje457 ' "Has vuelto a ser visible ya que no esta permitida la invisibilidad en este mapa." *-*  FontTypeNames.FONTTYPE_TALK
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("Has vuelto a ser visible ya que no esta permitida la invisibilidad en este mapa.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje458 ' "Has vuelto a ser visible ya que no esta permitido ocultarse en este mapa." *-*  FontTypeNames.FONTTYPE_TALK
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("Has vuelto a ser visible ya que no esta permitido ocultarse en este mapa.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje459 ' "¡Ocultarse no funciona aquí!" *-*  FontTypeNames.FONTTYPE_TALK
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("¡Ocultarse no funciona aquí!", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje460 ' "No hay un yacimiento de peces donde pescar." *-*  FontTypeNames.FONTTYPE_TALK
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("No hay un yacimiento de peces donde pescar.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje461 ' "No puedes pescar desde allí." *-*  FontTypeNames.FONTTYPE_TALK
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("No puedes pescar desde allí.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje462 ' "No puedes transportar dioses o admins." *-*  FontTypeNames.FONTTYPE_TALK
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("No puedes transportar dioses o admins.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje463 ' "Posición inválida." *-*  FontTypeNames.FONTTYPE_TALK
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("Posición inválida.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje464 ' "No puedes ver está información de un dios o administrador." *-*  FontTypeNames.FONTTYPE_TALK
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("No puedes ver está información de un dios o administrador.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje465 ' "Servidor.ini actualizado correctamente." *-*  FontTypeNames.FONTTYPE_TALK
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("Servidor.ini actualizado correctamente.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje466 ' "No puedes sumonear miembros del clan pretoriano de esta forma, utiliza /CrearClanPretoriano."  *-*   FontTypeNames.FONTTYPE_WARNING
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("No puedes sumonear miembros del clan pretoriano de esta forma, utiliza /CrearClanPretoriano.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje467 ' "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!" *-*  FontTypeNames.FONTTYPE_TALK
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje468 ' "¡¡¡No puedes robar a usuarios en consulta!!!" *-*  FontTypeNames.FONTTYPE_TALK
      With FontTypes(FontTypeNames.FONTTYPE_TALK)
         Call ShowConsoleMsg("¡¡¡No puedes robar a usuarios en consulta!!!", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje469 ' "¡¡Comercio cancelado, te están robando!!"  *-*   FontTypeNames.FONTTYPE_WARNING
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("¡¡Comercio cancelado, te están robando!!", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje470 ' "¡¡¡No puedes tirar este tipo de objeto!!!", FontTypeNames.FONTTYPE_FIGHT
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("¡¡¡No puedes tirar este tipo de objeto!!!", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje471 ' "No puedes vender este tipo de objeto.", FontTypeNames.FONTTYPE_INFO
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("No puedes vender este tipo de objeto.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje472 ' "Tu anillo rechaza los efectos del hechizo inmobilizar.", FontTypeNames.FONTTYPE_FIGHT
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Tu anillo rechaza los efectos del hechizo inmobilizar.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje473 ' "Tu anillo rechaza los efectos de la turbación.", FontTypeNames.FONTTYPE_FIGHT
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Tu anillo rechaza los efectos de la turbación.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje474 ' "Tu anillo rechaza los efectos de la ceguera.", FontTypeNames.FONTTYPE_FIGHT
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Tu anillo rechaza los efectos de la ceguera.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje475 ' "Tu anillo rechaza los efectos de la paralisis.", FontTypeNames.FONTTYPE_FIGHT
      With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
         Call ShowConsoleMsg("Tu anillo rechaza los efectos de la paralisis.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mansaje476 ' "El hechizo no pertenece a tu clase.", FontTypeNames.FONTTYPE_INFO
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El hechizo no pertenece a tu clase.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mansaje477 ' "El hechizo no pertenece a tu raza.", FontTypeNames.FONTTYPE_INFO
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("El hechizo no pertenece a tu raza.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje478     ' "Necesitas hacer click sobre un personaje.",  FontTypeNames.FONTTYPE_WARNING
      With FontTypes(FontTypeNames.FONTTYPE_WARNING)
         Call ShowConsoleMsg("Necesitas hacer click sobre un personaje.", .Red, .Green, .Blue, .bold, .italic)
      End With
    Case Is = eMensajes.Mensaje479 ' "Has conseguido algo de agua." *-*  FontTypeNames.FONTTYPE_INFO
      With FontTypes(FontTypeNames.FONTTYPE_INFO)
         Call ShowConsoleMsg("Has conseguido algo de agua.", .Red, .Green, .Blue, .bold, .italic)
      End With
      
End Select 'By TwIsT

End Sub

' Handles the StopWorking message.
Private Sub HandleStopWorking()
'***************************************************
'Author: Budi
'Last Modification: 12/01/09
'
'***************************************************

    Call incomingData.ReadByte
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg("¡Has terminado de trabajar!", .Red, .Green, .Blue, .bold, .italic)
    End With
    
    If frmMain.MacroTrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
End Sub


' Handles the HandleFormYesNo message.
Private Sub HandleFormYesNo()
'***************************************************
'Author: ^[GS]^
'Last Modification: 31/03/2013 - ^[GS]^
'
'***************************************************
On Error GoTo Fallo

    Call incomingData.ReadByte
    
    If bFormYesNo = True Then
        If frmFormYesNo.Visible = True Then
            Unload frmFormYesNo  ' Se cierra el formulario actual ya que es invalido de cualquier forma
        End If
        Exit Sub
    End If
    
    Dim DataString As String
    Dim FormType As Byte
    
    DataString = incomingData.ReadASCIIString
    FormType = incomingData.ReadByte
    
    'Set state and show form
    bFormYesNo = True
    
    ' Guardamos el tipo de petición para corroborar en el retorno
    frmFormYesNo.Tag = FormType
    
    ' Según el FormType cambia las opciones...
    Select Case FormType
        Case eAccionClick.Matrimonio ' Propuesta de Matrimonio
            frmFormYesNo.lMensaje.Caption = vbCr & DataString & " te está proponiendo matrimonio. ¿Aceptas?"
            frmFormYesNo.cAceptar.Tag = "Aceptar"
            frmFormYesNo.cRechazar.Caption = "Rechazar"
        Case eAccionClick.Divorcio
            frmFormYesNo.lMensaje.Caption = vbCr & DataString & " te está proponiendo divorciarte. ¿Aceptas?"
            frmFormYesNo.cAceptar.Tag = "Aceptar"
            frmFormYesNo.cRechazar.Caption = "Rechazar"
            
        Case Else
            frmFormYesNo.lMensaje.Caption = DataString
            frmFormYesNo.cAceptar.Tag = "Si"
            frmFormYesNo.cRechazar.Caption = "No"
    End Select
    
    frmFormYesNo.Show , frmMain

Fallo:
    ' Todos los sistemas experimentales deberían generar log
    Call LogError("HandleFormYesNo::Error " & Err.Number & " - " & Err.Description)
End Sub



' Handles the CancelOfferItem message.

Private Sub HandleCancelOfferItem()
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 05/03/10
'
'***************************************************
    Dim slot As Byte
    Dim amount As Long
    
    Call incomingData.ReadByte
    
    slot = incomingData.ReadByte
    
    With InvOfferComUsu(0)
        amount = .amount(slot)
        
        ' No tiene sentido que se quiten 0 unidades
        If amount <> 0 Then
            ' Actualizo el inventario general
            Call frmComerciarUsu.UpdateInvCom(.OBJIndex(slot), amount)
            
            ' Borro el item
            Call .SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
        End If
    End With
    
    ' Si era el único ítem de la oferta, no puede confirmarla
    If Not frmComerciarUsu.HasAnyItem(InvOfferComUsu(0)) And _
        Not frmComerciarUsu.HasAnyItem(InvOroComUsu(1)) Then Call frmComerciarUsu.HabilitarConfirmar(False)
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call frmComerciarUsu.PrintCommerceMsg("¡No puedes comerciar ese objeto!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 21 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    slot = Buffer.ReadByte()
    
    With UserBancoInventory(slot)
        .OBJIndex = Buffer.ReadInteger()
        .Name = Buffer.ReadASCIIString()
        .amount = Buffer.ReadInteger()
        .GrhIndex = Buffer.ReadInteger()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .MaxDef = Buffer.ReadInteger()
        .MinDef = Buffer.ReadInteger
        .Valor = Buffer.ReadLong()
        
        If Comerciando Then
            Call InvBanco(0).SetItem(slot, .OBJIndex, .amount, _
                .Equipped, .GrhIndex, .OBJType, .MaxHit, _
                .MinHit, .MaxDef, .MinDef, .Valor, .Name)
        End If
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ChangeSpellSlot message.

Private Sub HandleChangeSpellSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 09/07/2012 - ^[GS]^
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleChangeSpellSlot"
#End If

    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    slot = Buffer.ReadByte()
    
    UserHechizos(slot).Index = Buffer.ReadInteger()
    UserHechizos(slot).Name = Buffer.ReadASCIIString()
    UserHechizos(slot).GrhIndex = Buffer.ReadInteger()
    
    Spells.RenderSpells 'GDK: En el inventario el render se hace en el SetItem, como los spells no tiene setItem, habia que meter esto en algun lado jaja.
    
    'Debug.Print UserHechizos(slot).Name & " - " & UserHechizos(slot).GrhIndex
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the Attributes message.

Private Sub HandleAtributes()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 02/08/2012 - ^[GS]^
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleAtributes"
#End If
    
    If incomingData.length < 1 + NUMATRIBUTES Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = incomingData.ReadByte()
    Next i
    
    'Show them in character creation
    If EstadoLogin = E_MODO.Dados Then
        With frmCrearPersonaje
            If .Visible Then
                For i = 1 To NUMATRIBUTES
                    .lblAtributos(i).Caption = UserAtributos(i)
                Next i
                .UpdateStats
            End If
        End With
    Else
        LlegaronAtrib = True
    End If

End Sub

''
' Handles the BlacksmithWeapons message.

Private Sub HandleBlacksmithWeapons()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/08/2014 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim J As Long
    Dim K As Long
    
    Count = Buffer.ReadInteger()
    
    ReDim ArmasHerrero(Count) As tItemsConstruibles
    ReDim HerreroMejorar(0) As tItemsConstruibles
    
    For i = 1 To Count
        With ArmasHerrero(i)
            .Name = Buffer.ReadASCIIString()    'Get the object's name
            .GrhIndex = Buffer.ReadInteger()
            .LinH = Buffer.ReadInteger()        'The iron needed
            .LinP = Buffer.ReadInteger()        'The silver needed
            .LinO = Buffer.ReadInteger()        'The gold needed
            .OBJIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()
        End With
    Next i

    For i = 1 To MAX_LIST_ITEMS ' 0.13.3
        Set InvLingosHerreria(i) = New clsGraphicalInventory
    Next i

    With frmConstruirHerrero
        ' Inicializo los inventarios
        Call InvLingosHerreria(1).Initialize(DirectD3D8, .picLingotes0, 3, , , , , , False)
        Call InvLingosHerreria(2).Initialize(DirectD3D8, .picLingotes1, 3, , , , , , False)
        Call InvLingosHerreria(3).Initialize(DirectD3D8, .picLingotes2, 3, , , , , , False)
        Call InvLingosHerreria(4).Initialize(DirectD3D8, .picLingotes3, 3, , , , , , False)
        
        Call .HideExtraControls(Count)
        Call .RenderList(1, True)
    End With
    
    For i = 1 To Count
        With ArmasHerrero(i)
            If .Upgrade Then
                For K = 1 To Count
                    If .Upgrade = ArmasHerrero(K).OBJIndex Then
                        J = J + 1
                
                        ReDim Preserve HerreroMejorar(J) As tItemsConstruibles
                        
                        HerreroMejorar(J).Name = .Name
                        HerreroMejorar(J).GrhIndex = .GrhIndex
                        HerreroMejorar(J).OBJIndex = .OBJIndex
                        HerreroMejorar(J).UpgradeName = ArmasHerrero(K).Name
                        HerreroMejorar(J).UpgradeGrhIndex = ArmasHerrero(K).GrhIndex
                        HerreroMejorar(J).LinH = ArmasHerrero(K).LinH - .LinH * 0.85
                        HerreroMejorar(J).LinP = ArmasHerrero(K).LinP - .LinP * 0.85
                        HerreroMejorar(J).LinO = ArmasHerrero(K).LinO - .LinO * 0.85
                        
                        Exit For
                    End If
                Next K
            End If
        End With
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the BlacksmithArmors message.

Private Sub HandleBlacksmithArmors()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim J As Long
    Dim K As Long
    
    Count = Buffer.ReadInteger()
    
    ReDim ArmadurasHerrero(Count) As tItemsConstruibles
    
    For i = 1 To Count
        With ArmadurasHerrero(i)
            .Name = Buffer.ReadASCIIString()    'Get the object's name
            .GrhIndex = Buffer.ReadInteger()
            .LinH = Buffer.ReadInteger()        'The iron needed
            .LinP = Buffer.ReadInteger()        'The silver needed
            .LinO = Buffer.ReadInteger()        'The gold needed
            .OBJIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()
        End With
    Next i
    
    J = UBound(HerreroMejorar)
    
    For i = 1 To Count
        With ArmadurasHerrero(i)
            If .Upgrade Then
                For K = 1 To Count
                    If .Upgrade = ArmadurasHerrero(K).OBJIndex Then
                        J = J + 1
                
                        ReDim Preserve HerreroMejorar(J) As tItemsConstruibles
                        
                        HerreroMejorar(J).Name = .Name
                        HerreroMejorar(J).GrhIndex = .GrhIndex
                        HerreroMejorar(J).OBJIndex = .OBJIndex
                        HerreroMejorar(J).UpgradeName = ArmadurasHerrero(K).Name
                        HerreroMejorar(J).UpgradeGrhIndex = ArmadurasHerrero(K).GrhIndex
                        HerreroMejorar(J).LinH = ArmadurasHerrero(K).LinH - .LinH * 0.85
                        HerreroMejorar(J).LinP = ArmadurasHerrero(K).LinP - .LinP * 0.85
                        HerreroMejorar(J).LinO = ArmadurasHerrero(K).LinO - .LinO * 0.85
                        
                        Exit For
                    End If
                Next K
            End If
        End With
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the CarpenterObjects message.

Private Sub HandleCarpenterObjects()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/08/2014 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim J As Long
    Dim K As Long
    
    Count = Buffer.ReadInteger()
    
    ReDim ObjCarpintero(Count) As tItemsConstruibles
    ReDim CarpinteroMejorar(0) As tItemsConstruibles
    
    For i = 1 To Count
        With ObjCarpintero(i)
            .Name = Buffer.ReadASCIIString()        'Get the object's name
            .GrhIndex = Buffer.ReadInteger()
            .Madera = Buffer.ReadInteger()          'The wood needed
            .MaderaElfica = Buffer.ReadInteger()    'The elfic wood needed
            .OBJIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()
        End With
    Next i
    
    For i = 1 To MAX_LIST_ITEMS ' 0.13.3
        Set InvMaderasCarpinteria(i) = New clsGraphicalInventory
    Next i
    
    With frmConstruirCarp
        ' Inicializo los inventarios
        Call InvMaderasCarpinteria(1).Initialize(DirectD3D8, .picMaderas1, 2, , , , , , False)
        Call InvMaderasCarpinteria(2).Initialize(DirectD3D8, .picMaderas2, 2, , , , , , False)
        Call InvMaderasCarpinteria(3).Initialize(DirectD3D8, .picMaderas3, 2, , , , , , False)
        Call InvMaderasCarpinteria(4).Initialize(DirectD3D8, .picMaderas4, 2, , , , , , False)
        
        Call .HideExtraControls(Count)
        Call .RenderList(1)
    End With
    
    For i = 1 To Count
        With ObjCarpintero(i)
            If .Upgrade Then
                For K = 1 To Count
                    If .Upgrade = ObjCarpintero(K).OBJIndex Then
                        J = J + 1
                
                        ReDim Preserve CarpinteroMejorar(J) As tItemsConstruibles
                        
                        CarpinteroMejorar(J).Name = .Name
                        CarpinteroMejorar(J).GrhIndex = .GrhIndex
                        CarpinteroMejorar(J).OBJIndex = .OBJIndex
                        CarpinteroMejorar(J).UpgradeName = ObjCarpintero(K).Name
                        CarpinteroMejorar(J).UpgradeGrhIndex = ObjCarpintero(K).GrhIndex
                        CarpinteroMejorar(J).Madera = ObjCarpintero(K).Madera - .Madera * 0.85
                        CarpinteroMejorar(J).MaderaElfica = ObjCarpintero(K).MaderaElfica - .MaderaElfica * 0.85
                        
                        Exit For
                    End If
                Next K
            End If
        End With
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the RestOK message.

Private Sub HandleRestOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserDescansar = Not UserDescansar
End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 11/03/2012 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call MsgBox(Buffer.ReadASCIIString())
    
    If frmConnect.Visible And (Not frmCrearPersonaje.Visible) Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        ' Nos aseguramos que regrese al Connect!
        'frmconnect.
        frmConnect.Visible = True
        Call frmConnect.Show ' GSZAO
        Call frmConnect.EstadoSocket ' GSZAO
    End If

    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the Blind message.

Private Sub HandleBlind()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = True
End Sub

''
' Handles the Dumb message.

Private Sub HandleDumb()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = True
End Sub

''
' Handles the ShowSignal message.

Private Sub HandleShowSignal()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleShowSignal"
#End If

    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim tmp As String
    tmp = Buffer.ReadASCIIString()
    
    Call InitCartel(tmp, Buffer.ReadInteger())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 21 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    slot = Buffer.ReadByte()
    
    With NPCInventory(slot)
        .Name = Buffer.ReadASCIIString()
        .amount = Buffer.ReadInteger()
        .Valor = Buffer.ReadSingle()
        .GrhIndex = Buffer.ReadInteger()
        .OBJIndex = Buffer.ReadInteger()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .MaxDef = Buffer.ReadInteger()
        .MinDef = Buffer.ReadInteger
    End With
        
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 25/08/2012 - ^[GS]^
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleUpdateHungerAndThirst"
#End If
    
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMaxAGU = incomingData.ReadByte()
    UserMinAGU = incomingData.ReadByte()
    UserMaxHAM = incomingData.ReadByte()
    UserMinHAM = incomingData.ReadByte()
    
    frmMain.cStatSed.Max = UserMaxAGU
    frmMain.cStatHambre.Max = UserMaxHAM
    
    frmMain.cStatSed.Value = UserMinAGU
    frmMain.cStatHambre.Value = UserMinHAM
    
End Sub

''
' Handles the Fame message.

Private Sub HandleFame()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 29 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With UserReputacion
        .AsesinoRep = incomingData.ReadLong()
        .BandidoRep = incomingData.ReadLong()
        .BurguesRep = incomingData.ReadLong()
        .LadronesRep = incomingData.ReadLong()
        .NobleRep = incomingData.ReadLong()
        .PlebeRep = incomingData.ReadLong()
        .Promedio = incomingData.ReadLong()
    End With
    
    LlegoFama = True
End Sub

''
' Handles the MiniStats message.

Private Sub HandleMiniStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 20 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With UserEstadisticas
        .CiudadanosMatados = incomingData.ReadLong()
        .CriminalesMatados = incomingData.ReadLong()
        .UsuariosMatados = incomingData.ReadLong()
        .NpcsMatados = incomingData.ReadInteger()
        .Clase = ListaClases(incomingData.ReadByte())
        .PenaCarcel = incomingData.ReadLong()
    End With
End Sub

''
' Handles the LevelUp message.

Private Sub HandleLevelUp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleLevelUp"
#End If
    
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    SkillPoints = SkillPoints + incomingData.ReadInteger()
    
    Call frmMain.LightSkillStar(True)
End Sub

''
' Handles the AddForumMessage message.

Private Sub HandleAddForumMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim ForumType As eForumMsgType
    Dim Title As String
    Dim Message As String
    Dim Author As String
    Dim bAnuncio As Boolean
    Dim bSticky As Boolean
    
    ForumType = Buffer.ReadByte
    
    Title = Buffer.ReadASCIIString()
    Author = Buffer.ReadASCIIString()
    Message = Buffer.ReadASCIIString()
    
    If Not frmForo.ForoLimpio Then
        clsForos.ClearForums
        frmForo.ForoLimpio = True
    End If

    Call clsForos.AddPost(ForumAlignment(ForumType), Title, Author, Message, EsAnuncio(ForumType))
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowForumForm message.

Private Sub HandleShowForumForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmForo.Privilegios = incomingData.ReadByte
    frmForo.CanPostSticky = incomingData.ReadByte
    
    If Not MirandoForo Then
        frmForo.Show , frmMain
    End If
End Sub

''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleSetInvisible"
#End If
    
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    CharList(CharIndex).Invisible = incomingData.ReadBoolean()

End Sub

''
' Handles the DiceRoll message.

Private Sub HandleDiceRoll()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 02/08/2012 - ^[GS]^
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserAtributos(eAtributos.Fuerza) = incomingData.ReadByte()
    UserAtributos(eAtributos.Agilidad) = incomingData.ReadByte()
    UserAtributos(eAtributos.Inteligencia) = incomingData.ReadByte()
    UserAtributos(eAtributos.Carisma) = incomingData.ReadByte()
    UserAtributos(eAtributos.Constitucion) = incomingData.ReadByte()
    
    ' GSZAO
    CaptchaCode(0) = incomingData.ReadByte() Xor CaptchaKey
    CaptchaCode(1) = incomingData.ReadByte() Xor CaptchaKey
    CaptchaCode(2) = incomingData.ReadByte() Xor CaptchaKey
    CaptchaCode(3) = incomingData.ReadByte() Xor CaptchaKey
    
    With frmCrearPersonaje
        .lblAtributos(eAtributos.Fuerza) = UserAtributos(eAtributos.Fuerza)
        .lblAtributos(eAtributos.Agilidad) = UserAtributos(eAtributos.Agilidad)
        .lblAtributos(eAtributos.Inteligencia) = UserAtributos(eAtributos.Inteligencia)
        .lblAtributos(eAtributos.Carisma) = UserAtributos(eAtributos.Carisma)
        .lblAtributos(eAtributos.Constitucion) = UserAtributos(eAtributos.Constitucion)
                
        .UpdateStats
        If Not .Visible Then ' GSZAO
            frmConnect.Visible = True
            Call Audio.PlayMIDI("7.mid")
            .Show vbModal
            DoEvents
        End If
    End With
End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMeditar = Not UserMeditar
End Sub

''
' Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = False
End Sub

''
' Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = False
End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 30/10/2012 - ^[GS]^
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleSendSkills"
#End If

    If incomingData.length < (1 + NUMSKILLS * 2) Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = incomingData.ReadByte()
        PorcentajeSkills(i) = incomingData.ReadByte()
    Next i
    
    LlegaronSkills = True
    
End Sub

''
' Handles the TrainerCreatureList message.

Private Sub HandleTrainerCreatureList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim creatures() As String
    Dim i As Long
    
    creatures = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(creatures())
        Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
    Next i
    frmEntrenador.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the GuildNews message.

Private Sub HandleGuildNews()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'11/19/09: Pato - Is optional show the frmGuildNews form
'***************************************************
    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim guildList() As String
    Dim i As Long
    Dim sTemp As String
    
    'Get news' string
    frmGuildNews.news = Buffer.ReadASCIIString()
    
    'Get Enemy guilds list
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(guildList)
        sTemp = frmGuildNews.txtClanesGuerra.Text
        frmGuildNews.txtClanesGuerra.Text = sTemp & guildList(i) & vbCrLf
    Next i
    
    'Get Allied guilds list
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(guildList)
        sTemp = frmGuildNews.txtClanesAliados.Text
        frmGuildNews.txtClanesAliados.Text = sTemp & guildList(i) & vbCrLf
    Next i
    
    If ClientAOSetup.bGuildNews Or bShowGuildNews Then frmGuildNews.Show vbModeless, frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the OfferDetails message.

Private Sub HandleOfferDetails()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the User Die
Private Sub HandleDieAlocate()
'***************************************************
'Author: Standelf
'Last Modification: 10/12/2012
'
'***************************************************
On Error GoTo ErrHandler
    Call incomingData.ReadByte
    ABD = 255
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0

    If error <> 0 Then _
        Err.Raise error
End Sub




''
' Handles the Online message.

Private Sub HandleOnline()
'***************************************************
'Author: ^[GS]^
'Last Modification: 14/05/2013 - ^[GS]^
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleOnline"
#End If

On Error GoTo ErrHandler

    Call incomingData.ReadByte
    frmMain.lblOnline.Caption = incomingData.ReadInteger ' GSZAO
    If incomingData.ReadBoolean = True Then ' GSZAO
        Call AddtoRichTextBox(frmMain.RecTxt, "Usuarios conectados: " & frmMain.lblOnline.Caption, 255, 0, 0, True, False, True)
    End If

ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the AlianceProposalsList message.

Private Sub HandleAlianceProposalsList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim vsGuildList() As String
    Dim i As Long
    
    vsGuildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    Call frmPeaceProp.lista.Clear
    For i = 0 To UBound(vsGuildList())
        Call frmPeaceProp.lista.AddItem(vsGuildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the PeaceProposalsList message.

Private Sub HandlePeaceProposalsList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim guildList() As String
    Dim i As Long
    
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    Call frmPeaceProp.lista.Clear
    For i = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.PAZ
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the CharacterInfo message.

Private Sub HandleCharacterInfo()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 11/08/2012 - ^[GS]^
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleCharacterInfo"
#End If
    
    If incomingData.length < 35 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmCharInfo
        If .frmType = CharInfoFrmType.frmMembers Then
            .cRechazar.Visible = False
            .cAceptar.Visible = False
            .cEchar.Visible = True
            .cPeticion.Visible = False
        Else
            .cRechazar.Visible = True
            .cAceptar.Visible = True
            .cEchar.Visible = False
            .cPeticion.Visible = True
        End If
        
        .Nombre.Caption = Buffer.ReadASCIIString()
        .Raza.Caption = ListaRazas(Buffer.ReadByte())
        .Clase.Caption = ListaClases(Buffer.ReadByte())
        
        If Buffer.ReadByte() = 1 Then
            .Genero.Caption = "Hombre"
        Else
            .Genero.Caption = "Mujer"
        End If
        
        .Nivel.Caption = Buffer.ReadByte()
        .Oro.Caption = Buffer.ReadLong()
        .Banco.Caption = Buffer.ReadLong()
        
        Dim reputation As Long
        reputation = Buffer.ReadLong()
        
        .reputacion.Caption = reputation
        
        .txtPeticiones.Text = Buffer.ReadASCIIString()
        .guildactual.Caption = Buffer.ReadASCIIString()
        .txtMiembro.Text = Buffer.ReadASCIIString()
        
        Dim armada As Boolean
        Dim caos As Boolean
        
        armada = Buffer.ReadBoolean()
        caos = Buffer.ReadBoolean()
        
        If armada Then
            .ejercito.Caption = "Armada Real"
        ElseIf caos Then
            .ejercito.Caption = "Legión Oscura"
        End If
        
        .Ciudadanos.Caption = CStr(Buffer.ReadLong())
        .criminales.Caption = CStr(Buffer.ReadLong())
        
        If reputation > 0 Then
            .status.Caption = " Ciudadano"
            .status.ForeColor = vbBlue
        Else
            .status.Caption = " Criminal"
            .status.ForeColor = vbRed
        End If
        
        Call .Show(vbModeless, frmMain)
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the GuildLeaderInfo message.

Private Sub HandleGuildLeaderInfo()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim i As Long
    Dim List() As String
    
    With frmGuildLeader
        'Get list of existing guilds
        GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .guildslist.Clear
        
        For i = 0 To UBound(GuildNames())
            Call .guildslist.AddItem(GuildNames(i))
        Next i
        
        'Get list of guild's members
        GuildMembers = Split(Buffer.ReadASCIIString(), SEPARATOR)
        .Miembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .members.Clear
        
        For i = 0 To UBound(GuildMembers())
            Call .members.AddItem(GuildMembers(i))
        Next i
        
        .txtguildnews = Buffer.ReadASCIIString()
        
        'Get list of join requests
        List = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .solicitudes.Clear
        
        For i = 0 To UBound(List())
            Call .solicitudes.AddItem(List(i))
        Next i
        
        .FillLogo (Buffer.ReadASCIIString)
        
        .Show , frmMain
    End With

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the GuildDetails message.

Private Sub HandleGuildDetails()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 11/08/2012 - ^[GS]^
'***************************************************
    If incomingData.length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmGuildBrief
        .cDeclararGuerra.Visible = .EsLeader
        .cOfrecerAlianza.Visible = .EsLeader
        .cOfrecerPaz.Visible = .EsLeader
        
        .Nombre.Caption = Buffer.ReadASCIIString()
        .fundador.Caption = Buffer.ReadASCIIString()
        .creacion.Caption = Buffer.ReadASCIIString()
        .lider.Caption = Buffer.ReadASCIIString()
        .web.Caption = Buffer.ReadASCIIString()
        .Miembros.Caption = Buffer.ReadInteger()
        
        If Buffer.ReadBoolean() Then
            .eleccion.Caption = "ABIERTA"
        Else
            .eleccion.Caption = "CERRADA"
        End If
        
        .lblAlineacion.Caption = Buffer.ReadASCIIString()
        .Enemigos.Caption = Buffer.ReadInteger()
        .Aliados.Caption = Buffer.ReadInteger()
        .antifaccion.Caption = Buffer.ReadASCIIString()
        
        Dim codexStr() As String
        Dim i As Long
        
        codexStr = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        For i = 0 To 7
            .Codex(i).Caption = codexStr(i)
        Next i
        
        .Desc.Text = Buffer.ReadASCIIString()
        .FillLogo Buffer.ReadASCIIString()
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
    frmGuildBrief.Show vbModeless, frmMain
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowGuildAlign message.

Private Sub HandleShowGuildAlign()
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmEligeAlineacion.Show vbModeless, frmMain
End Sub

''
' Handles the ShowGuildFundationForm message.

Private Sub HandleShowGuildFundationForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    CreandoClan = True
    frmGuildFoundation.Show , frmMain
End Sub

''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserParalizado = Not UserParalizado
End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString())
    Call frmUserRequest.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the TradeOK message.

Private Sub HandleTradeOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If frmComerciar.Visible Then
        Dim i As Long
        
        'Update user inventory
        For i = 1 To MAX_INVENTORY_SLOTS
            ' Agrego o quito un item en su totalidad
            If Inventario.OBJIndex(i) <> InvComUsu.OBJIndex(i) Then
                With Inventario
                    Call InvComUsu.SetItem(i, .OBJIndex(i), _
                    .amount(i), .Equipped(i), .GrhIndex(i), _
                    .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                    .Valor(i), .ItemName(i))
                End With
            ' Vendio o compro cierta cantidad de un item que ya tenia
            ElseIf Inventario.amount(i) <> InvComUsu.amount(i) Then
                Call InvComUsu.ChangeSlotItemAmount(i, Inventario.amount(i))
            End If
        Next i
        
        ' Fill Npc inventory
        For i = 1 To 20
            ' Compraron la totalidad de un item, o vendieron un item que el npc no tenia
            If NPCInventory(i).OBJIndex <> InvComNpc.OBJIndex(i) Then
                With NPCInventory(i)
                    Call InvComNpc.SetItem(i, .OBJIndex, _
                    .amount, 0, .GrhIndex, _
                    .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                    .Valor, .Name)
                End With
            ' Compraron o vendieron cierta cantidad (no su totalidad)
            ElseIf NPCInventory(i).amount <> InvComNpc.amount(i) Then
                Call InvComNpc.ChangeSlotItemAmount(i, NPCInventory(i).amount)
            End If
        Next i
    
    End If
End Sub

''
' Handles the BankOK message.

Private Sub HandleBankOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    If frmBancoObj.Visible Then
        
        For i = 1 To Inventario.MaxObjs
            With Inventario
                Call InvBanco(1).SetItem(i, .OBJIndex(i), .amount(i), _
                    .Equipped(i), .GrhIndex(i), .OBJType(i), .MaxHit(i), _
                    .MinHit(i), .MaxDef(i), .MinDef(i), .Valor(i), .ItemName(i))
            End With
        Next i
        
        'Alter order according to if we bought or sold so the labels and grh remain the same
        If frmBancoObj.LasActionBuy Then
            'frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
            'frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
        Else
            'frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
            'frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
        End If
        
        frmBancoObj.NoPuedeMover = False
    End If
       
End Sub

''
' Handles the ChangeUserTradeSlot message.

Private Sub HandleChangeUserTradeSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 22 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    Dim OfferSlot As Byte
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    OfferSlot = Buffer.ReadByte
    
    With Buffer
        If OfferSlot = GOLD_OFFER_SLOT Then
            Call InvOroComUsu(2).SetItem(1, .ReadInteger(), .ReadLong(), 0, _
                                            .ReadInteger(), .ReadByte(), .ReadInteger(), _
                                            .ReadInteger(), .ReadInteger(), .ReadInteger(), .ReadLong(), .ReadASCIIString())
        Else
            Call InvOfferComUsu(1).SetItem(OfferSlot, .ReadInteger(), .ReadLong(), 0, _
                                            .ReadInteger(), .ReadByte(), .ReadInteger(), _
                                            .ReadInteger(), .ReadInteger(), .ReadInteger(), .ReadLong(), .ReadASCIIString())
        End If
    End With
    
    Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & " ha modificado su oferta.", FontTypeNames.FONTTYPE_VENENO)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ClientConfig message.

Private Sub HandleClientConfig()
'***************************************************
'Author: ^[GS]^
'Last Modification: 07/04/2012 (maTih.-)
'                   Ahora el cliente recibe el MeditarRapido.
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleClientConfig"
#End If
    
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

    CfgDiaNoche = incomingData.ReadBoolean() ' DiaNoche
    CfgSistemaLuces = incomingData.ReadBoolean() ' Sistema de Luces
    CfgSiempreNombres = incomingData.ReadBoolean()  ' Mostrar Siempre los Nombres
    ClMeditarRapido = incomingData.ReadBoolean() 'MeditarRapido?
    
End Sub

Private Sub HandleCreateParticleInChar()
'***************************************************
'Author: maTih.-
'Last Modification: -
'***************************************************

Exit Sub 'NOTA: Deshabilitado hasta proximo aviso! 06/07/2012

Dim CharIndex       As Integer
Dim OtherCharIndex  As Integer
Dim ParticleID      As Integer
Dim TargetX         As Long
Dim TargetY         As Long

'Remove PacketID.
Call incomingData.ReadByte

CharIndex = incomingData.ReadInteger()
OtherCharIndex = incomingData.ReadInteger()
ParticleID = incomingData.ReadInteger()

TargetX = Engine_TPtoSPX(CharList(CharIndex).Pos.X)
TargetY = Engine_TPtoSPY(CharList(CharIndex).Pos.Y)
' HAY UN ERROR EN EL CENTRADO X-Y
'Debug.Print "Particulas X" & TargetX & " - Y" & TargetY

'Mismo charIndex? es una meditación/warp.
If Not OtherCharIndex <> CharIndex Then

    'Destruir partícula?
    If Not ParticleID <> -1 Then
        'El char tenia partícula?
        If CharList(CharIndex).ParticleIndex <> 0 Then
            'Al que se le ocurra algo mejor que me avise ! jaja
            effect(CharList(CharIndex).ParticleIndex).Used = False
        End If
    Else 'crear.
        'Warp o meditación?
        If ParticleID <> EffectNum_Summon Then
            CharList(CharIndex).ParticleIndex = Effect_Meditate_Begin(TargetX, TargetY, 1, 150, 10, 1000)
        Else    'warp
            CharList(CharIndex).ParticleIndex = Effect_Summon_Begin(TargetX, TargetY, 1, 150)
        End If
    End If
    
End If

End Sub

''
' Handles the SpawnList message.

Private Sub HandleSpawnList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim creatureList() As String
    Dim i As Long
    
    creatureList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(creatureList())
        Call frmSpawnList.lstCriaturas.AddItem(creatureList(i))
    Next i
    frmSpawnList.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowSOSForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim sosList() As String
    Dim i As Long
    
    sosList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(sosList())
        Call frmMSG.List1.AddItem(sosList(i))
    Next i
    
    frmMSG.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowDenounces message.

Private Sub HandleShowDenounces() ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleShowDenounces"
#End If
    
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim DenounceList() As String
    Dim DenounceIndex As Long
    
    DenounceList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        For DenounceIndex = 0 To UBound(DenounceList())
            Call AddtoRichTextBox(frmMain.RecTxt, DenounceList(DenounceIndex), .Red, .Green, .Blue, .bold, .italic)
        Next DenounceIndex
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub



''
' Handles the ShowSOSForm message.

Private Sub HandleShowPartyForm()
'***************************************************
'Author: Budi
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim members() As String
    Dim i As Long
    
    EsPartyLeader = CBool(Buffer.ReadByte())
       
    members = Split(Buffer.ReadASCIIString(), SEPARATOR)
    For i = 0 To UBound(members())
        Call frmParty.lstMembers.AddItem(members(i))
    Next i
    
    frmParty.lblTotalExp.Caption = Buffer.ReadLong
    frmParty.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub



''
' Handles the ShowMOTDEditionForm message.

Private Sub HandleShowMOTDEditionForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'*************************************Su**************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    frmCambiaMotd.txtMotd.Text = Buffer.ReadASCIIString()
    frmCambiaMotd.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowGMPanelForm message.

Private Sub HandleShowGMPanelForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmPanelGm.Show vbModeless, frmMain
End Sub

''
' Handles the UserNameList message.

Private Sub HandleUserNameList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim userList() As String
    Dim i As Long
    
    userList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    If frmPanelGm.Visible Then
        frmPanelGm.cboListaUsus.Clear
        For i = 0 To UBound(userList())
            Call frmPanelGm.cboListaUsus.AddItem(userList(i))
        Next i
        If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the Pong message.

Private Sub HandlePong()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, "El ping es " & (GetTickCount - pingTime) & " ms.", 255, 0, 0, True, False, True)
    
    pingTime = 0
End Sub



Private Sub HandleGuildMemberInfo()
'***************************************************
'Author: ZaMa
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmGuildMember
        'Clear guild's list
        .lstClanes.Clear
        
        GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        Dim i As Long
        For i = 0 To UBound(GuildNames())
            Call .lstClanes.AddItem(GuildNames(i))
        Next i
        
        'Get list of guild's members
        GuildMembers = Split(Buffer.ReadASCIIString(), SEPARATOR)
        .lblCantMiembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .lstMiembros.Clear
        
        For i = 0 To UBound(GuildMembers())
            Call .lstMiembros.AddItem(GuildMembers(i))
        Next i
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
        
        .Show vbModeless, frmMain
    End With
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 30/10/2012 - ^[GS]^
'
'***************************************************
#If Testeo = 1 Then
    Debug.Print Now & " - IN: HandleUpdateTagAndStatus"
#End If
    
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim NickColor As Byte
    Dim UserTag As String
    
    CharIndex = Buffer.ReadInteger()
    NickColor = Buffer.ReadByte()
    UserTag = Buffer.ReadASCIIString()
    
    'Update char status adn tag!
    With CharList(CharIndex)
        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0
        End If
        
        ' GSZAO
        If (NickColor And eNickColor.ieNewbie) <> 0 Then
            .Newbie = 1
        Else
            .Newbie = 0
        End If
        ' GSZAO
        
        ' GSZAO
        If (NickColor And eNickColor.ieMuerto) = True Then
            .muerto = True
        Else
            .muerto = False
        End If
        ' GSZAO
        
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        
        .Nombre = UserTag
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandleQuestDetails() ' GSZAO
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Recibe y maneja el paquete QuestDetails del servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If incomingData.length < 15 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo ErrHandler
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
   
    Dim tmpStr As String
    Dim tmpByte As Byte
    Dim QuestEmpezada As Boolean
    Dim i As Integer
   
    With Buffer
        'Leemos el id del paquete
        Call .ReadByte
       
        'Nos fijamos si se trata de una quest empezada, para poder leer los NPCs que se han matado.
        QuestEmpezada = IIf(.ReadByte, True, False)
       
        tmpStr = "Misión: " & .ReadASCIIString & vbCrLf
        tmpStr = tmpStr & "Detalles: " & .ReadASCIIString & vbCrLf
        tmpStr = tmpStr & "Nivel requerido: " & .ReadByte & vbCrLf
       
        tmpStr = tmpStr & vbCrLf & "OBJETIVOS" & vbCrLf
       
        tmpByte = .ReadByte
        If tmpByte Then 'Hay NPCs
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) Matar " & .ReadInteger & " " & .ReadASCIIString & "."
                If QuestEmpezada Then
                    tmpStr = tmpStr & " (Has matado " & .ReadInteger & ")" & vbCrLf
                Else
                    tmpStr = tmpStr & vbCrLf
                End If
            Next i
        End If
       
        tmpByte = .ReadByte
        If tmpByte Then 'Hay OBJs
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) Conseguir " & .ReadInteger & " " & .ReadASCIIString & "." & vbCrLf
            Next i
        End If
 
        tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf
        tmpStr = tmpStr & "*) Oro: " & .ReadLong & " monedas de oro." & vbCrLf
        tmpStr = tmpStr & "*) Experiencia: " & .ReadLong & " puntos de experiencia." & vbCrLf
       
        tmpByte = .ReadByte
        If tmpByte Then
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) " & .ReadInteger & " " & .ReadASCIIString & vbCrLf
            Next i
        End If
    End With
   
    'Determinamos que formulario se muestra, según si recibimos la información y la quest está empezada o no.
    If QuestEmpezada Then
        frmQuests.txtInfo.Text = tmpStr
    Else
        frmQuestInfo.txtInfo.Text = tmpStr
        frmQuestInfo.Show vbModeless, frmMain
    End If
   
    Call incomingData.CopyBuffer(Buffer)
   
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
   
    'Destroy auxiliar buffer
    Set Buffer = Nothing
 
    If error <> 0 Then _
        Err.Raise error
End Sub
 
Public Sub HandleQuestListSend() ' GSZAO
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Recibe y maneja el paquete QuestListSend del servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If incomingData.length < 1 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo ErrHandler
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
   
    Dim i As Integer
    Dim tmpByte As Byte
    Dim tmpStr As String
   
    'Leemos el id del paquete
    Call Buffer.ReadByte
     
    'Leemos la cantidad de quests que tiene el usuario
    tmpByte = Buffer.ReadByte
   
    'Limpiamos el ListBox y el TextBox del formulario
    frmQuests.lstQuests.Clear
    frmQuests.txtInfo.Text = vbNullString
       
    'Si el usuario tiene quests entonces hacemos el handle
    If tmpByte Then
        'Leemos el string
        tmpStr = Buffer.ReadASCIIString
       
        'Agregamos los items
        For i = 1 To tmpByte
            frmQuests.lstQuests.AddItem ReadField(i, tmpStr, 45)
        Next i
    End If
   
    'Mostramos el formulario
    frmQuests.Show vbModeless, frmMain
   
    'Pedimos la información de la primer quest (si la hay)
    If tmpByte Then Call modProtocol.WriteQuestDetailsRequest(1)
   
    'Copiamos de vuelta el buffer
    Call incomingData.CopyBuffer(Buffer)
 
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
   
    'Destroy auxiliar buffer
    Set Buffer = Nothing
 
    If error <> 0 Then _
        Err.Raise error
End Sub


''
' Writes the "LoginExistingChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginExistingChar()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/06/2012 - ^[GS]^
'Writes the "LoginExistingChar" message to the outgoing data buffer
'***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginExistingChar)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(UserPassword)
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString(SEncriptar(GetSerialHD())) ' GSZAO
        
    End With
    
#If Testeo = 1 Then
    Call LogTesteo("El usuario " & UserName & " intenta conectarse.")
#End If
End Sub

''
' Writes the "ThrowDices" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteThrowDices()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 02/08/2012 - ^[GS]^
'Writes the "ThrowDices" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ThrowDices)
    Call outgoingData.WriteByte(CaptchaKey)   ' GSZAO
End Sub

''
' Writes the "LoginNewChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginNewChar()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/06/2012 - ^[GS]^
'Writes the "LoginNewChar" message to the outgoing data buffer
'***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginNewChar)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(UserPassword)
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString(SEncriptar(GetSerialHD())) ' GSZAO
        Call .WriteByte(UserRaza)
        Call .WriteByte(UserSexo)
        Call .WriteByte(UserClase)
        Call .WriteInteger(UserHead)
        
        Call .WriteASCIIString(UserEmail)
        
        Call .WriteByte(UserHogar)
    End With
End Sub

''
' Writes the "Talk" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalk(ByVal Chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Talk" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Talk)
        
        Call .WriteASCIIString(Chat)
    End With
End Sub

''
' Writes the "Yell" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteYell(ByVal Chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Yell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Yell)
        
        Call .WriteASCIIString(Chat)
    End With
End Sub

''
' Writes the "Whisper" message to the outgoing data buffer.
'
' @param    charIndex The index of the char to whom to whisper.
' @param    chat The chat text to be sent to the user.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhisper(ByVal CharName As String, ByVal Chat As String) ' 0.13.3
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/11/2011 - ^[GS]^
'Writes the "Whisper" message to the outgoing data buffer
'03/12/10: Enanoh - Ahora se envía el nick y no el charindex.
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Whisper)
        
        Call .WriteASCIIString(CharName)
        
        Call .WriteASCIIString(Chat)
    End With
End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal Heading As E_Heading)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/09/2012 - ^[GS]^
'Writes the "Walk" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Walk)
        Call .WriteByte(Heading)
    End With
End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestPositionUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestPositionUpdate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)
End Sub

''
' Writes the "Attack" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttack()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Attack" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Attack)
End Sub

''
' Writes the "PickUp" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePickUp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PickUp" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PickUp)
End Sub

''
' Writes the "SafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SafeToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.SafeToggle)
End Sub

''
' Writes the "ResuscitationSafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationToggle()
'**************************************************************
'Author: Rapsodius
'Creation Date: 10/10/07
'Writes the Resuscitation safe toggle packet to the outgoing data buffer.
'**************************************************************
    Call outgoingData.WriteByte(ClientPacketID.ResuscitationSafeToggle)
End Sub

''
' Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestGuildLeaderInfo()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestGuildLeaderInfo)
End Sub

Public Sub WriteRequestFormYesNo(ByVal bAccion As Byte, ByVal bResp As Byte)
'***************************************************
'Author: ^[GS]^
'Last Modification: 18/03/2013 - ^[GS]^
'Writes the "RequestFormYesNo" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestFormYesNo)
    Call outgoingData.WriteByte(bAccion)
    Call outgoingData.WriteByte(bResp)

End Sub


Public Sub WriteRequestPartyForm()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "RequestPartyForm" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestPartyForm)

End Sub

''
' Writes the "ItemUpgrade" message to the outgoing data buffer.
'
' @param    ItemIndex The index to the item to upgrade.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemUpgrade(ByVal ItemIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 12/09/09
'Writes the "ItemUpgrade" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ItemUpgrade)
    Call outgoingData.WriteInteger(ItemIndex)
End Sub

''
' Writes the "RequestAtributes" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAtributes()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestAtributes" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestAtributes)
End Sub

''
' Writes the "RequestFame" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestFame()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestFame" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestFame)
End Sub

''
' Writes the "RequestSkills" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestSkills()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestSkills" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestSkills)
End Sub

''
' Writes the "RequestMiniStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMiniStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestMiniStats" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestMiniStats)
End Sub

''
' Writes the "CommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceEnd" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CommerceEnd)
End Sub

''
' Writes the "UserCommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceEnd" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceEnd)
End Sub

''
' Writes the "UserCommerceConfirm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceConfirm()
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UserCommerceConfirm" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceConfirm)
End Sub

''
' Writes the "BankEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankEnd" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BankEnd)
End Sub

''
' Writes the "UserCommerceOk" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOk()
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/10/07
'Writes the "UserCommerceOk" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceOk)
End Sub

''
' Writes the "UserCommerceReject" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceReject()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceReject" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceReject)
End Sub

''
' Writes the "Drop" message to the outgoing data buffer.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDrop(ByVal slot As Byte, ByVal amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Drop" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Drop)
        
        Call .WriteByte(slot)
        Call .WriteInteger(amount)
    End With
End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CastSpell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CastSpell)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LeftClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "DoubleClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoubleClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DoubleClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.DoubleClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "Work" message to the outgoing data buffer.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWork(ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Work" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Work)
        
        Call .WriteByte(Skill)
    End With
End Sub

''
' Writes the "UseSpellMacro" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseSpellMacro()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UseSpellMacro" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UseSpellMacro)
End Sub

''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseItem(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UseItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UseItem)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "CraftBlacksmith" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftBlacksmith(ByVal Item As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CraftBlacksmith" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftBlacksmith)
        
        Call .WriteInteger(Item)
    End With
End Sub

''
' Writes the "CraftCarpenter" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftCarpenter(ByVal Item As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CraftCarpenter" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftCarpenter)
        
        Call .WriteInteger(Item)
    End With
End Sub

''
' Writes the "ShowGuildNews" message to the outgoing data buffer.
'

Public Sub WriteShowGuildNews()
'***************************************************
'Author: ZaMa
'Last Modification: 21/02/2010
'Writes the "ShowGuildNews" message to the outgoing data buffer
'***************************************************
 
     outgoingData.WriteByte (ClientPacketID.ShowGuildNews)
End Sub


''
' Writes the "WorkLeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkLeftClick(ByVal X As Byte, ByVal Y As Byte, ByVal Skill As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 18/03/2013 - ^[GS]^
'Writes the "WorkLeftClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.WorkLeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .WriteByte(Skill)
    End With
End Sub

''
' Writes the "CreateNewGuild" message to the outgoing data buffer.
'
' @param    desc    The guild's description
' @param    name    The guild's name
' @param    site    The guild's website
' @param    codex   Array of all rules of the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNewGuild(ByVal Desc As String, ByVal Name As String, ByVal Site As String, ByRef Codex() As String, ByVal rLogo As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNewGuild" message to the outgoing data buffer
'***************************************************
    Dim temp As String
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNewGuild)
        
        Call .WriteASCIIString(Desc)
        Call .WriteASCIIString(Name)
        Call .WriteASCIIString(Site)
        
        For i = LBound(Codex()) To UBound(Codex())
            temp = temp & Codex(i) & SEPARATOR
        Next i
        
        If Len(temp) Then _
            temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)
        
        Call .WriteASCIIString(rLogo)
    End With
End Sub

''
' Writes the "SpellInfo" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpellInfo(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpellInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SpellInfo)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "EquipItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.EquipItem)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "ChangeHeading" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeHeading(ByVal Heading As E_Heading)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeHeading" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeHeading)
        
        Call .WriteByte(Heading)
    End With
End Sub

''
' Writes the "ModifySkills" message to the outgoing data buffer.
'
' @param    skillEdt a-based array containing for each skill the number of points to add to it.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteModifySkills(ByRef skillEdt() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ModifySkills" message to the outgoing data buffer
'***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ModifySkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(skillEdt(i))
        Next i
    End With
End Sub

''
' Writes the "Train" message to the outgoing data buffer.
'
' @param    creature Position within the list provided by the server of the creature to train against.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrain(ByVal creature As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Train" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Train)
        
        Call .WriteByte(creature)
    End With
End Sub

''
' Writes the "CommerceBuy" message to the outgoing data buffer.
'
' @param    slot Position within the NPC's inventory in which the desired item is.
' @param    amount Number of items to buy.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceBuy(ByVal slot As Byte, ByVal amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceBuy" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceBuy)
        
        Call .WriteByte(slot)
        Call .WriteInteger(amount)
    End With
End Sub

''
' Writes the "BankExtractItem" message to the outgoing data buffer.
'
' @param    slot Position within the bank in which the desired item is.
' @param    amount Number of items to extract.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractItem(ByVal slot As Byte, ByVal amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankExtractItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractItem)
        
        Call .WriteByte(slot)
        Call .WriteInteger(amount)
    End With
End Sub

''
' Writes the "CommerceSell" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to sell.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceSell(ByVal slot As Byte, ByVal amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceSell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceSell)
        
        Call .WriteByte(slot)
        Call .WriteInteger(amount)
    End With
End Sub

''
' Writes the "BankDeposit" message to the outgoing data buffer.
'
' @param    slot Position within the user inventory in which the desired item is.
' @param    amount Number of items to deposit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDeposit(ByVal slot As Byte, ByVal amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankDeposit" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDeposit)
        
        Call .WriteByte(slot)
        Call .WriteInteger(amount)
    End With
End Sub

''
' Writes the "ForumPost" message to the outgoing data buffer.
'
' @param    title The message's title.
' @param    message The body of the message.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForumPost(ByVal Title As String, ByVal Message As String, ByVal ForumMsgType As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForumPost" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForumPost)
        
        Call .WriteByte(ForumMsgType)
        Call .WriteASCIIString(Title)
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "MoveSpell" message to the outgoing data buffer.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MoveSpell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveSpell)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "MoveBank" message to the outgoing data buffer.
'
' @param    upwards True if the item will be moved up in the list, False if it will be moved downwards.
' @param    slot Bank List slot where the item which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveBank(ByVal upwards As Boolean, ByVal slot As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 06/14/09
'Writes the "MoveBank" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveBank)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "ClanCodexUpdate" message to the outgoing data buffer.
'
' @param    desc New description of the clan.
' @param    codex New codex of the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteClanCodexUpdate(ByVal Desc As String, ByRef Codex() As String, ByVal rLogo As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ClanCodexUpdate" message to the outgoing data buffer
'***************************************************
    Dim temp As String
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ClanCodexUpdate)
        
        Call .WriteASCIIString(Desc)
        
        For i = LBound(Codex()) To UBound(Codex())
            temp = temp & Codex(i) & SEPARATOR
        Next i
        
        If Len(temp) Then _
            temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)
        Call .WriteASCIIString(rLogo)
    End With
End Sub

''
' Writes the "UserCommerceOffer" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to offer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOffer(ByVal slot As Byte, ByVal amount As Long, ByVal OfferSlot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceOffer" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UserCommerceOffer)
        
        Call .WriteByte(slot)
        Call .WriteLong(amount)
        Call .WriteByte(OfferSlot)
    End With
End Sub

Public Sub WriteCommerceChat(ByVal Chat As String)
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'Writes the "CommerceChat" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceChat)
        
        Call .WriteASCIIString(Chat)
    End With
End Sub


''
' Writes the "GuildAcceptPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptPeace(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAcceptPeace" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptPeace)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildRejectAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectAlliance(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRejectAlliance" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectAlliance)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildRejectPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectPeace(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRejectPeace" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectPeace)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildAcceptAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptAlliance(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAcceptAlliance" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptAlliance)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildOfferPeace" message to the outgoing data buffer.
'
' @param    guild The guild to whom peace is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferPeace(ByVal guild As String, ByVal proposal As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOfferPeace" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferPeace)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)
    End With
End Sub

''
' Writes the "GuildOfferAlliance" message to the outgoing data buffer.
'
' @param    guild The guild to whom an aliance is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferAlliance(ByVal guild As String, ByVal proposal As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOfferAlliance" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferAlliance)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)
    End With
End Sub

''
' Writes the "GuildAllianceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAllianceDetails(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAllianceDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAllianceDetails)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildPeaceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose peace proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeaceDetails(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildPeaceDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildPeaceDetails)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer.
'
' @param    username The user who wants to join the guild whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestJoinerInfo)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildAlliancePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAlliancePropList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAlliancePropList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildAlliancePropList)
End Sub

''
' Writes the "GuildPeacePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeacePropList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildPeacePropList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildPeacePropList)
End Sub

''
' Writes the "GuildDeclareWar" message to the outgoing data buffer.
'
' @param    guild The guild to which to declare war.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDeclareWar(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildDeclareWar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildDeclareWar)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildNewWebsite" message to the outgoing data buffer.
'
' @param    url The guild's new website's URL.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNewWebsite(ByVal URL As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildNewWebsite" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildNewWebsite)
        
        Call .WriteASCIIString(URL)
    End With
End Sub

''
' Writes the "GuildAcceptNewMember" message to the outgoing data buffer.
'
' @param    username The name of the accepted player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAcceptNewMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptNewMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildRejectNewMember" message to the outgoing data buffer.
'
' @param    username The name of the rejected player.
' @param    reason The reason for which the player was rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal Reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRejectNewMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectNewMember)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(Reason)
    End With
End Sub

''
' Writes the "GuildKickMember" message to the outgoing data buffer.
'
' @param    username The name of the kicked player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildKickMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildKickMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildKickMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildUpdateNews" message to the outgoing data buffer.
'
' @param    news The news to be posted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildUpdateNews(ByVal news As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildUpdateNews" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildUpdateNews)
        
        Call .WriteASCIIString(news)
    End With
End Sub

''
' Writes the "GuildMemberInfo" message to the outgoing data buffer.
'
' @param    username The user whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildMemberInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMemberInfo)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildOpenElections" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOpenElections()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOpenElections" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildOpenElections)
End Sub

''
' Writes the "GuildRequestMembership" message to the outgoing data buffer.
'
' @param    guild The guild to which to request membership.
' @param    application The user's application sheet.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestMembership" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestMembership)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(Application)
    End With
End Sub

''
' Writes the "GuildRequestDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestDetails(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestDetails)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "Online" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnline()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Online" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Online)
End Sub

''
' Writes the "Quit" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteQuit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/16/08
'Writes the "Quit" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Quit)
End Sub

''
' Writes the "GuildLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeave()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildLeave" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildLeave)
End Sub

''
' Writes the "RequestAccountState" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAccountState()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestAccountState" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestAccountState)
End Sub

''
' Writes the "PetStand" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetStand()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PetStand" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PetStand)
End Sub

''
' Writes the "PetFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetFollow()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PetFollow" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PetFollow)
End Sub

''
' Writes the "ReleasePet" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReleasePet()
'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2009
'Writes the "ReleasePet" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ReleasePet)
End Sub


''
' Writes the "TrainList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TrainList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.TrainList)
End Sub

''
' Writes the "Rest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Rest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Rest)
End Sub

''
' Writes the "Meditate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Meditate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Meditate)
End Sub

''
' Writes the "Resucitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResucitate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Resucitate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Resucitate)
End Sub

''
' Writes the "Consultation" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsultation() ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 23/11/2011 - ^[GS]^
'Writes the "Consultation" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Consultation)

End Sub

''
' Writes the "Heal" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHeal()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Heal" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Heal)
End Sub

''
' Writes the "Help" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHelp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Help" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Help)
End Sub

''
' Writes the "RequestStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestStats" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestStats)
End Sub

''
' Writes the "CommerceStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceStart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceStart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CommerceStart)
End Sub

''
' Writes the "BankStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankStart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankStart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BankStart)
End Sub

''
' Writes the "Enlist" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnlist()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Enlist" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Enlist)
End Sub

''
' Writes the "Information" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInformation()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Information" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Information)
End Sub

''
' Writes the "Reward" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReward()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Reward" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Reward)
End Sub

''
' Writes the "RequestMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMOTD()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestMOTD" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestMOTD)
End Sub

''
' Writes the "UpTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpTime()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpTime" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Uptime)
End Sub

''
' Writes the "PartyLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyLeave()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyLeave" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyLeave)
End Sub

''
' Writes the "PartyCreate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyCreate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyCreate)
End Sub

''
' Writes the "PartyJoin" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyJoin()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyJoin" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyJoin)
End Sub

''
' Writes the "Inquiry" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiry()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Inquiry" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Inquiry)
End Sub

''
' Writes the "GuildMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "PartyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CentinelReport" message to the outgoing data buffer.
'
' @param    number The number to report to the centinel.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCentinelReport(ByVal Number As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CentinelReport" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CentinelReport)
        
        Call .WriteInteger(Number)
    End With
End Sub

''
' Writes the "GuildOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnline()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOnline" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildOnline)
End Sub

''
' Writes the "PartyOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyOnline()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyOnline" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyOnline)
End Sub

''
' Writes the "CouncilMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the other council members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CouncilMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CouncilMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "RoleMasterRequest" message to the outgoing data buffer.
'
' @param    message The message to send to the role masters.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoleMasterRequest(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoleMasterRequest" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RoleMasterRequest)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "GMRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMRequest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMRequest)
End Sub

''
' Writes the "BugReport" message to the outgoing data buffer.
'
' @param    message The message explaining the reported bug.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBugReport(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BugReport" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.bugReport)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ChangeDescription" message to the outgoing data buffer.
'
' @param    desc The new description of the user's character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeDescription(ByVal Desc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeDescription" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeDescription)
        
        Call .WriteASCIIString(Desc)
    End With
End Sub

''
' Writes the "GuildVote" message to the outgoing data buffer.
'
' @param    username The user to vote for clan leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildVote(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildVote" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildVote)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Punishments" message to the outgoing data buffer.
'
' @param    username The user whose's  punishments are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePunishments(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Punishments" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Punishments)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ChangePassword" message to the outgoing data buffer.
'
' @param    oldPass Previous password.
' @param    newPass New password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangePassword(ByRef oldPass As String, ByRef newPass As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/10/07
'Last Modified By: Rapsodius
'Writes the "ChangePassword" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangePassword)
        Call .WriteASCIIString(oldPass)
        Call .WriteASCIIString(newPass)
    End With
End Sub

''
' Writes the "Gamble" message to the outgoing data buffer.
'
' @param    amount The amount to gamble.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGamble(ByVal amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Gamble" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Gamble)
        
        Call .WriteInteger(amount)
    End With
End Sub

''
' Writes the "InquiryVote" message to the outgoing data buffer.
'
' @param    opt The chosen option to vote for.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiryVote(ByVal opt As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "InquiryVote" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.InquiryVote)
        
        Call .WriteByte(opt)
    End With
End Sub

''
' Writes the "LeaveFaction" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeaveFaction()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LeaveFaction" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.LeaveFaction)
End Sub

''
' Writes the "BankExtractGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to extract from the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractGold(ByVal amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankExtractGold" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractGold)
        
        Call .WriteLong(amount)
    End With
End Sub

''
' Writes the "BankDepositGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to deposit in the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDepositGold(ByVal amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankDepositGold" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDepositGold)
        
        Call .WriteLong(amount)
    End With
End Sub

''
' Writes the "Denounce" message to the outgoing data buffer.
'
' @param    message The message to send with the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDenounce(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Denounce" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Denounce)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "GuildFundate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/21/2001
'Writes the "GuildFundate" message to the outgoing data buffer
'14/12/2009: ZaMa - Now first checks if the user can foundate a guild.
'03/21/2001: Pato - Deleted de clanType param.
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildFundate)
End Sub

''
' Writes the "GuildFundation" message to the outgoing data buffer.
'
' @param    clanType The alignment of the clan to be founded.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundation(ByVal clanType As eClanType)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "GuildFundation" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildFundation)
        
        Call .WriteByte(clanType)
    End With
End Sub

''
' Writes the "PartyKick" message to the outgoing data buffer.
'
' @param    username The user to kick fro mthe party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyKick)
            
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "PartySetLeader" message to the outgoing data buffer.
'
' @param    username The user to set as the party's leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartySetLeader(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartySetLeader" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartySetLeader)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "PartyAcceptMember" message to the outgoing data buffer.
'
' @param    username The user to accept into the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyAcceptMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyAcceptMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyAcceptMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildMemberList" message to the outgoing data buffer.
'
' @param    guild The guild whose member list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberList(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildMemberList" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildMemberList)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "InitCrafting" message to the outgoing data buffer.
'
' @param    Cantidad The final aumont of item to craft.
' @param    NroPorCiclo The amount of items to craft per cicle.

Public Sub WriteInitCrafting(ByVal cantidad As Long, ByVal NroPorCiclo As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 29/01/2010
'Writes the "InitCrafting" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.InitCrafting)
        Call .WriteLong(cantidad)
        
        Call .WriteInteger(NroPorCiclo)
    End With
End Sub

''
' Writes the "Home" message to the outgoing data buffer.
'
Public Sub WriteHome()
'***************************************************
'Author: Budi
'Last Modification: 01/06/10
'Writes the "Home" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Home)
    End With
End Sub



''
' Writes the "GMMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to the other GMs online.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GMMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ShowName" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowName()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowName" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.showName)
End Sub

''
' Writes the "OnlineRoyalArmy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineRoyalArmy()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineRoyalArmy" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineRoyalArmy)
End Sub

''
' Writes the "OnlineChaosLegion" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineChaosLegion()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineChaosLegion" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineChaosLegion)
End Sub

''
' Writes the "GoNearby" message to the outgoing data buffer.
'
' @param    username The suer to approach.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoNearby(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GoNearby" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GoNearby)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Comment" message to the outgoing data buffer.
'
' @param    message The message to leave in the log as a comment.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteComment(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Comment" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Comment)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ServerTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerTime()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerTime" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.serverTime)
End Sub

''
' Writes the "Where" message to the outgoing data buffer.
'
' @param    username The user whose position is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhere(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Where" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Where)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "CreaturesInMap" message to the outgoing data buffer.
'
' @param    map The map in which to check for the existing creatures.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreaturesInMap(ByVal Map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreaturesInMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreaturesInMap)
        
        Call .WriteInteger(Map)
    End With
End Sub

''
' Writes the "WarpMeToTarget" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpMeToTarget()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarpMeToTarget" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.WarpMeToTarget)
End Sub

''
' Writes the "WarpChar" message to the outgoing data buffer.
'
' @param    username The user to be warped. "YO" represent's the user's char.
' @param    map The map to which to warp the character.
' @param    x The x position in the map to which to waro the character.
' @param    y The y position in the map to which to waro the character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpChar(ByVal UserName As String, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarpChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.WarpChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "Silence" message to the outgoing data buffer.
'
' @param    username The user to silence.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSilence(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Silence" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Silence)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "SOSShowList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSShowList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SOSShowList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SOSShowList)
End Sub

''
' Writes the "SOSRemove" message to the outgoing data buffer.
'
' @param    username The user whose SOS call has been already attended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSRemove(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SOSRemove" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SOSRemove)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GoToChar" message to the outgoing data buffer.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoToChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GoToChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GoToChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "invisible" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvisible()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "invisible" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Invisible)
End Sub

''
' Writes the "GMPanel" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMPanel()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMPanel" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.GMPanel)
End Sub

''
' Writes the "RequestUserList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestUserList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestUserList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RequestUserList)
End Sub

''
' Writes the "Working" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorking()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Working" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Working)
End Sub

''
' Writes the "Hiding" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHiding()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Hiding" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Hiding)
End Sub

''
' Writes the "Jail" message to the outgoing data buffer.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteJail(ByVal UserName As String, ByVal Reason As String, ByVal Time As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Jail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Jail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(Reason)
        
        Call .WriteByte(Time)
    End With
End Sub

''
' Writes the "KillNPC" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPC()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillNPC" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillNPC)
End Sub

''
' Writes the "WarnUser" message to the outgoing data buffer.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarnUser(ByVal UserName As String, ByVal Reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarnUser" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.WarnUser)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(Reason)
    End With
End Sub

''
' Writes the "EditChar" message to the outgoing data buffer.
'
' @param    UserName    The user to be edited.
' @param    editOption  Indicates what to edit in the char.
' @param    arg1        Additional argument 1. Contents depend on editoption.
' @param    arg2        Additional argument 2. Contents depend on editoption.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEditChar(ByVal UserName As String, ByVal EditOption As eEditOptions, ByVal arg1 As String, ByVal arg2 As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "EditChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.EditChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteByte(EditOption)
        
        Call .WriteASCIIString(arg1)
        Call .WriteASCIIString(arg2)
    End With
End Sub

''
' Writes the "RequestCharInfo" message to the outgoing data buffer.
'
' @param    username The user whose information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharInfo)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharStats" message to the outgoing data buffer.
'
' @param    username The user whose stats are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharStats(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharStats" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharStats)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharGold" message to the outgoing data buffer.
'
' @param    username The user whose gold is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharGold(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharGold" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharGold)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub
    
''
' Writes the "RequestCharInventory" message to the outgoing data buffer.
'
' @param    username The user whose inventory is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInventory(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharInventory" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharInventory)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharBank" message to the outgoing data buffer.
'
' @param    username The user whose banking information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharBank(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharBank" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharBank)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharSkills" message to the outgoing data buffer.
'
' @param    username The user whose skills are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharSkills(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharSkills" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharSkills)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReviveChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReviveChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ReviveChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "OnlineGM" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineGM()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineGM" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineGM)
End Sub

''
' Writes the "OnlineMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineMap(ByVal Map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'Writes the "OnlineMap" message to the outgoing data buffer
'26/03/2009: Now you don't need to be in the map to use the comand, so you send the map to server
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.OnlineMap)
        
        Call .WriteInteger(Map)
    End With
End Sub

''
' Writes the "Forgive" message to the outgoing data buffer.
'
' @param    username The user to be forgiven.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForgive(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Forgive" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Forgive)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Kick" message to the outgoing data buffer.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Kick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Kick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Execute" message to the outgoing data buffer.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExecute(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Execute" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Execute)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "BanChar" message to the outgoing data buffer.
'
' @param    username The user to be banned.
' @param    reason The reson for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanChar(ByVal UserName As String, ByVal Reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BanChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.banChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteASCIIString(Reason)
    End With
End Sub

''
' Writes the "UnbanChar" message to the outgoing data buffer.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UnbanChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnbanChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "NPCFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCFollow()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCFollow" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.NPCFollow)
End Sub

''
' Writes the "SummonChar" message to the outgoing data buffer.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSummonChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SummonChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SummonChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "SpawnListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnListRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnListRequest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SpawnListRequest)
End Sub

''
' Writes the "SpawnCreature" message to the outgoing data buffer.
'
' @param    creatureIndex The index of the creature in the spawn list to be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnCreature" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SpawnCreature)
        
        Call .WriteInteger(creatureIndex)
    End With
End Sub

''
' Writes the "ResetNPCInventory" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetNPCInventory()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetNPCInventory" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ResetNPCInventory)
End Sub

''
' Writes the "CleanWorld" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanWorld()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CleanWorld" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.CleanWorld)
End Sub

''
' Writes the "ServerMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ServerMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "MapMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMapMessage(ByVal Message As String) ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'Writes the "MapMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MapMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "NickToIP" message to the outgoing data buffer.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNickToIP(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NickToIP" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.nickToIP)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "IPToNick" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIPToNick(ByRef Ip() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "IPToNick" message to the outgoing data buffer
'***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.IPToNick)
        
        For i = LBound(Ip()) To UBound(Ip())
            Call .WriteByte(Ip(i))
        Next i
    End With
End Sub

''
' Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
' @param    guild The guild whose online player list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnlineMembers(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOnlineMembers" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildOnlineMembers)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "TeleportCreate" message to the outgoing data buffer.
'
' @param    map the map to which the teleport will lead.
' @param    x The position in the x axis to which the teleport will lead.
' @param    y The position in the y axis to which the teleport will lead.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportCreate(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal Radio As Byte = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TeleportCreate" message to the outgoing data buffer
'***************************************************
    With outgoingData
            Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TeleportCreate)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .WriteByte(Radio)
    End With
End Sub

''
' Writes the "TeleportDestroy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportDestroy()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TeleportDestroy" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TeleportDestroy)
End Sub

''
' Writes the "RainToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RainToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RainToggle)
End Sub

''
' Writes the "SetCharDescription" message to the outgoing data buffer.
'
' @param    desc The description to set to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetCharDescription(ByVal Desc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetCharDescription" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetCharDescription)
        
        Call .WriteASCIIString(Desc)
    End With
End Sub

''
' Writes the "ForceMIDIToMap" message to the outgoing data buffer.
'
' @param    midiID The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal Map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceMIDIToMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMIDIToMap)
        
        Call .WriteByte(midiID)
        
        Call .WriteInteger(Map)
    End With
End Sub

''
' Writes the "ForceWAVEToMap" message to the outgoing data buffer.
'
' @param    waveID  The ID of the wave file to play.
' @param    Map     The map into which to play the given wave.
' @param    x       The position in the x axis in which to play the given wave.
' @param    y       The position in the y axis in which to play the given wave.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceWAVEToMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceWAVEToMap)
        
        Call .WriteByte(waveID)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "RoyalArmyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoyalArmyMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RoyalArmyMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ChaosLegionMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the chaos legion member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosLegionMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosLegionMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CitizenMessage" message to the outgoing data buffer.
'
' @param    message The message to send to citizens.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCitizenMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CitizenMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CitizenMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CriminalMessage" message to the outgoing data buffer.
'
' @param    message The message to send to criminals.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCriminalMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CriminalMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CriminalMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "TalkAsNPC" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalkAsNPC(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TalkAsNPC" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TalkAsNPC)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "DestroyAllItemsInArea" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyAllItemsInArea()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DestroyAllItemsInArea" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DestroyAllItemsInArea)
End Sub

''
' Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted into the royal army council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AcceptRoyalCouncilMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted as a chaos council member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AcceptChaosCouncilMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ItemsInTheFloor" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemsInTheFloor()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ItemsInTheFloor" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ItemsInTheFloor)
End Sub

''
' Writes the "MakeDumb" message to the outgoing data buffer.
'
' @param    username The name of the user to be made dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumb(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MakeDumb" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MakeDumb)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
' @param    username The name of the user who will no longer be dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumbNoMore(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MakeDumbNoMore" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MakeDumbNoMore)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "DumpIPTables" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumpIPTables()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumpIPTables" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.dumpIPTables)
End Sub

''
' Writes the "CouncilKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CouncilKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CouncilKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "SetTrigger" message to the outgoing data buffer.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetTrigger" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetTrigger)
        
        Call .WriteByte(Trigger)
    End With
End Sub

''
' Writes the "AskTrigger" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAskTrigger()
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'Writes the "AskTrigger" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.AskTrigger)
End Sub

''
' Writes the "BannedIPList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BannedIPList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.BannedIPList)
End Sub

''
' Writes the "BannedIPReload" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPReload()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BannedIPReload" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.BannedIPReload)
End Sub

''
' Writes the "GuildBan" message to the outgoing data buffer.
'
' @param    guild The guild whose members will be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildBan(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildBan" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildBan)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "CheckHD" message to the outgoing data buffer.
'
'@param   username The name of the user to be checked.
Public Sub WriteVerHD(ByVal UserName As String)
'***************************************************
'Author: ArzenaTh
'Last Modification: 01/09/10
'Checkeamos la HD del usuario.
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.VerHD)
       
        Call .WriteASCIIString(UserName)
    End With
End Sub
 
''
' Writes the "UnBanHD" message to the outgoing data buffer
'
'@param    username The name of the user to be unbanned.
Public Sub WriteUnBanHD(ByVal HD As String)
'***************************************************
'Author: ArzenaTh
'Last Modification: 01/09/10
'Unbaneamos al usuario con su HD baneado.
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnBanHD)
       
        Call .WriteASCIIString(HD)
    End With
End Sub
 
''
' Writes the "BanHD" message to the outgoing data buffer
'
'@param    username The name of the user to be banned.
Public Sub WriteBanHD(ByVal UserName As String)
'***************************************************
'Author: ArzenaTh
'Last Modification: 01/09/10
'Baneamos la HD del usuario.
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.BanHD)
       
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "BanIP" message to the outgoing data buffer.
'
' @param    byIp    If set to true, we are banning by IP, otherwise the ip of a given character.
' @param    IP      The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @param    nick    The nick of the player whose ip will be banned.
' @param    reason  The reason for the ban.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanIP(ByVal byIp As Boolean, ByRef Ip() As Byte, ByVal Nick As String, ByVal Reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BanIP" message to the outgoing data buffer
'***************************************************
    If byIp And UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.BanIP)
        
        Call .WriteBoolean(byIp)
        
        If byIp Then
            For i = LBound(Ip()) To UBound(Ip())
                Call .WriteByte(Ip(i))
            Next i
        Else
            Call .WriteASCIIString(Nick)
        End If
        
        Call .WriteASCIIString(Reason)
    End With
End Sub

''
' Writes the "UnbanIP" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanIP(ByRef Ip() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UnbanIP" message to the outgoing data buffer
'***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnbanIP)
        
        For i = LBound(Ip()) To UBound(Ip())
            Call .WriteByte(Ip(i))
        Next i
    End With
End Sub

''
' Writes the "CreateItem" message to the outgoing data buffer.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateItem(ByVal ItemIndex As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateItem)
        Call .WriteInteger(ItemIndex)
    End With
End Sub

''
' Writes the "DestroyItems" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyItems()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DestroyItems" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DestroyItems)
End Sub

''
' Writes the "ChaosLegionKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Chaos Legion.
' @param    reason The reson for which the user is kicked from the Chaos Legion.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionKick(ByVal UserName As String, ByVal Reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 27/07/2012 - ^[GS]^
'Writes the "ChaosLegionKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosLegionKick)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(Reason) ' 0.13.5
    End With
End Sub

''
' Writes the "RoyalArmyKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Royal Army.
' @param    reason The reson for which the user is kicked from the Royal Army.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyKick(ByVal UserName As String, ByVal Reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 27/07/2012 - ^[GS]^
'Writes the "RoyalArmyKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RoyalArmyKick)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(Reason) ' 0.13.5
    End With
End Sub

''
' Writes the "ForceMIDIAll" message to the outgoing data buffer.
'
' @param    midiID The id of the midi file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIAll(ByVal midiID As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceMIDIAll" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMIDIAll)
        
        Call .WriteByte(midiID)
    End With
End Sub

''
' Writes the "ForceWAVEAll" message to the outgoing data buffer.
'
' @param    waveID The id of the wave file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceWAVEAll" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceWAVEAll)
        
        Call .WriteByte(waveID)
    End With
End Sub

''
' Writes the "RemovePunishment" message to the outgoing data buffer.
'
' @param    username The user whose punishments will be altered.
' @param    punishment The id of the punishment to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemovePunishment(ByVal UserName As String, ByVal punishment As Byte, ByVal NewText As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemovePunishment" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RemovePunishment)
        
        Call .WriteASCIIString(UserName)
        Call .WriteByte(punishment)
        Call .WriteASCIIString(NewText)
    End With
End Sub

''
' Writes the "TileBlockedToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTileBlockedToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TileBlockedToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TileBlockedToggle)
End Sub

''
' Writes the "KillNPCNoRespawn" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPCNoRespawn()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillNPCNoRespawn" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillNPCNoRespawn)
End Sub

''
' Writes the "KillAllNearbyNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillAllNearbyNPCs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillAllNearbyNPCs" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillAllNearbyNPCs)
End Sub

''
' Writes the "LastIP" message to the outgoing data buffer.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLastIP(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LastIP" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.LastIP)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ChangeMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMOTD()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMOTD" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ChangeMOTD)
End Sub

''
' Writes the "SetMOTD" message to the outgoing data buffer.
'
' @param    message The message to be set as the new MOTD.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetMOTD(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetMOTD" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetMOTD)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "SystemMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to all players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSystemMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SystemMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SystemMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CreateNPC" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPC(ByVal NPCIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNPC" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateNPC)
        
        Call .WriteInteger(NPCIndex)
    End With
End Sub

''
' Writes the "CreateNPCWithRespawn" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPCWithRespawn(ByVal NPCIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNPCWithRespawn" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateNPCWithRespawn)
        
        Call .WriteInteger(NPCIndex)
    End With
End Sub

''
' Writes the "ImperialArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of imperial armour to be altered.
' @param    objectIndex The index of the new object to be set as the imperial armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImperialArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ImperialArmour" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ImperialArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)
    End With
End Sub

''
' Writes the "ChaosArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of chaos armour to be altered.
' @param    objectIndex The index of the new object to be set as the chaos armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosArmour" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)
    End With
End Sub

''
' Writes the "NavigateToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.NavigateToggle)
End Sub

''
' Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerOpenToUsersToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ServerOpenToUsersToggle)
End Sub

''
' Writes the "TurnOffServer" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnOffServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TurnOffServer" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TurnOffServer)
End Sub

''
' Writes the "TurnCriminal" message to the outgoing data buffer.
'
' @param    username The name of the user to turn into criminal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnCriminal(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TurnCriminal" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TurnCriminal)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ResetFactions" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any faction.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetFactions(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetFactions" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ResetFactions)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharFromGuild" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RemoveCharFromGuild)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharMail(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharMail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharMail)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "AlterPassword" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    copyFrom The name of the user from which to copy the password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterPassword(ByVal UserName As String, ByVal CopyFrom As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterPassword" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterPassword)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(CopyFrom)
    End With
End Sub

''
' Writes the "AlterMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newMail The new email of the player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterMail(ByVal UserName As String, ByVal newMail As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterMail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterMail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newMail)
    End With
End Sub

''
' Writes the "AlterName" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newName The new user name.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterName" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterName)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newName)
    End With
End Sub

''
' Writes the "ToggleCentinelActivated" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteToggleCentinelActivated()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ToggleCentinelActivated" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ToggleCentinelActivated)
End Sub

''
' Writes the "DoBackup" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoBackup()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DoBackup" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DoBackUp)
End Sub

''
' Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
' @param    guild The guild to listen to.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildMessages(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGuildMessages" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ShowGuildMessages)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "SaveMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveMap()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SaveMap" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SaveMap)
End Sub

''
' Writes the "ChangeMapInfoPK" message to the outgoing data buffer.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMapInfoPK" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoPK)
        
        Call .WriteBoolean(isPK)
    End With
End Sub

''
' Writes the "WriteChangeMapInfoNoOcultar" message to the outgoing data buffer.
'
' @param    PermitirOcultar True if the map permits to hide, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoOcultar(ByVal PermitirOcultar As Boolean) ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 11/03/2012
'Writes the "WriteChangeMapInfoNoOcultar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoOcultar)
        
        Call .WriteBoolean(PermitirOcultar)
    End With
End Sub

''
' Writes the "ChangeMapInfoNoInvocar" message to the outgoing data buffer.
'
' @param    PermitirInvocar True if the map permits to invoke, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvocar(ByVal PermitirInvocar As Boolean) ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 18/09/2010
'Writes the "ChangeMapInfoNoInvocar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoInvocar)
        
        Call .WriteBoolean(PermitirInvocar)
    End With
End Sub


''
' Writes the "ChangeMapInfoBackup" message to the outgoing data buffer.
'
' @param    backup True if the map is to be backuped, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMapInfoBackup" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoBackup)
        
        Call .WriteBoolean(backup)
    End With
End Sub

''
' Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer.
'
' @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoRestricted)
        
        Call .WriteASCIIString(restrict)
    End With
End Sub

''
' Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer.
'
' @param    nomagic TRUE if no magic is to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoMagic)
        
        Call .WriteBoolean(nomagic)
    End With
End Sub

''
' Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer.
'
' @param    noinvi TRUE if invisibility is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoInvi)
        
        Call .WriteBoolean(noinvi)
    End With
End Sub
                            
''
' Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer.
'
' @param    noresu TRUE if resurection is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoResu)
        
        Call .WriteBoolean(noresu)
    End With
End Sub
                        
''
' Writes the "ChangeMapInfoLand" message to the outgoing data buffer.
'
' @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoLand(ByVal land As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoLand" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoLand)
        
        Call .WriteASCIIString(land)
    End With
End Sub
                        
''
' Writes the "ChangeMapInfoZone" message to the outgoing data buffer.
'
' @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoZone(ByVal zone As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoZone" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoZone)
        
        Call .WriteASCIIString(zone)
    End With
End Sub

''
' Writes the "ChangeMapInfoStealNpc" message to the outgoing data buffer.
'
' @param    forbid TRUE if stealNpc forbiden.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoStealNpc(ByVal forbid As Boolean) ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 23/11/2011 - ^[GS]^
'Writes the "ChangeMapInfoStealNpc" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoStealNpc)
        
        Call .WriteBoolean(forbid)
    End With
End Sub

''
' Writes the "SaveChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveChars()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SaveChars" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SaveChars)
End Sub

''
' Writes the "CleanSOS" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanSOS()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CleanSOS" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.CleanSOS)
End Sub

''
' Writes the "ShowServerForm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowServerForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowServerForm" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ShowServerForm)
End Sub

''
' Writes the "ShowDenouncesList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowDenouncesList() ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 23/11/2011 - ^[GS]^
'Writes the "ShowDenouncesList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ShowDenouncesList)
End Sub

''
' Writes the "EnableDenounces" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnableDenounces() ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 23/11/2011 - ^[GS]^
'Writes the "EnableDenounces" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.EnableDenounces)
End Sub


''
' Writes the "KickAllChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKickAllChars()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KickAllChars" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KickAllChars)
End Sub

''
' Writes the "ReloadNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadNPCs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadNPCs" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadNPCs)
End Sub

''
' Writes the "ReloadServerIni" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadServerIni()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadServerIni" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadServerIni)
End Sub

''
' Writes the "ReloadSpells" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadSpells()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadSpells" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadSpells)
End Sub

''
' Writes the "ReloadObjects" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadObjects()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadObjects" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadObjects)
End Sub

''
' Writes the "Restart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Restart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Restart)
End Sub

''
' Writes the "ResetAutoUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetAutoUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetAutoUpdate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ResetAutoUpdate)
End Sub

''
' Writes the "ChatColor" message to the outgoing data buffer.
'
' @param    r The red component of the new chat color.
' @param    g The green component of the new chat color.
' @param    b The blue component of the new chat color.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatColor(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChatColor" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChatColor)
        
        Call .WriteByte(r)
        Call .WriteByte(g)
        Call .WriteByte(b)
    End With
End Sub

''
' Writes the "Ignored" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIgnored()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Ignored" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Ignored)
End Sub

''
' Writes the "CheckSlot" message to the outgoing data buffer.
'
' @param    UserName    The name of the char whose slot will be checked.
' @param    slot        The slot to be checked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCheckSlot(ByVal UserName As String, ByVal slot As Byte)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "CheckSlot" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CheckSlot)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "Ping" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/01/2007
'Writes the "Ping" message to the outgoing data buffer
'***************************************************
    'Prevent the timer from being cut
    If pingTime <> 0 Then Exit Sub
    
    Call outgoingData.WriteByte(ClientPacketID.Ping)
    
    ' Avoid computing errors due to frame rate
    Call FlushBuffer
    DoEvents
    
    pingTime = GetTickCount
End Sub

''
' Writes the "ShareNpc" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShareNpc()
'***************************************************
'Author: ZaMa
'Last Modification: 15/04/2010
'Writes the "ShareNpc" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ShareNpc)
End Sub

''
' Writes the "StopSharingNpc" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteStopSharingNpc()
'***************************************************
'Author: ZaMa
'Last Modification: 15/04/2010
'Writes the "StopSharingNpc" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.StopSharingNpc)
End Sub

''
' Writes the "SetIniVar" message to the outgoing data buffer.
'
' @param    sLlave the name of the key which contains the value to edit
' @param    sClave the name of the value to edit
' @param    sValor the new value to set to sClave
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetIniVar(ByRef sLlave As String, ByRef sClave As String, ByRef sValor As String)
'***************************************************
'Author: Brian Chaia (BrianPr)
'Last Modification: 21/06/2009
'Writes the "SetIniVar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetIniVar)
        
        Call .WriteASCIIString(sLlave)
        Call .WriteASCIIString(sClave)
        Call .WriteASCIIString(sValor)
    End With
End Sub

''
' Writes the "CreatePretorianClan" message to the outgoing data buffer.
'
' @param    Map         The map in which create the pretorian clan.
' @param    X           The x pos where the king is settled.
' @param    Y           The y pos where the king is settled.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreatePretorianClan(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte) ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 23/11/2011 - ^[GS]^
'Writes the "CreatePretorianClan" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreatePretorianClan)
        Call .WriteInteger(Map)
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "DeletePretorianClan" message to the outgoing data buffer.
'
' @param    Map         The map which contains the pretorian clan to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDeletePretorianClan(ByVal Map As Integer) ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 23/11/2011 - ^[GS]^
'Writes the "DeletePretorianClan" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RemovePretorianClan)
        Call .WriteInteger(Map)
    End With
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String
    
    With outgoingData
        If .length = 0 Then _
            Exit Sub
        
        sndData = .ReadASCIIStringFixed(.length)
        
        Call SendData(sndData)
    End With
End Sub

''
' Sends the data using the socket controls in the MainForm.
'
' @param    sdData  The data to be sent to the server.

Private Sub SendData(ByRef sdData As String)
    
    'No enviamos nada si no estamos conectados
    If Not frmMain.Socket1.IsWritable Then
        'Put data back in the bytequeue
        Call outgoingData.WriteASCIIStringFixed(sdData)
        
        Exit Sub
    End If
    
    If Not frmMain.Socket1.Connected Then Exit Sub
    
    'Send data!
    Call frmMain.Socket1.Write(sdData, Len(sdData))

End Sub


''
' Writes the "AdminCargos" message to the outgoing data buffer.
'
' @param    Cargo       The charge.
' @param    Accion      The action.
' @param    UserName    The user to be edited. (only if is needed)
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAdminCargos(ByVal Cargo As eCargos, ByVal Accion As eAcciones, ByVal UserName As String)
'***************************************************
'Author: ^[GS]^
'Last Modification: 19/06/2011
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AdminCargos)
        
        Call .WriteByte(Cargo)
        Call .WriteByte(Accion)
        
        If Accion <> eAcciones.a_Listar Then ' nos ahorramos ancho de banda enviando un string que no se usa...
            Call .WriteASCIIString(UserName)
        End If
        
    End With
End Sub

''
' Writes the "MapMessage" message to the outgoing data buffer.
'
' @param    Dialog The new dialog of the NPC.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetDialog(ByVal dialog As String)
'***************************************************
'Author: Amraphen
'Last Modification: 18/11/2010
'Writes the "SetDialog" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetDialog)
        
        Call .WriteASCIIString(dialog)
    End With
End Sub

''
' Writes the "Impersonate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImpersonate() ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 23/11/2011 - ^[GS]^
'Writes the "Impersonate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Impersonate)
End Sub

''
' Writes the "Imitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImitate() ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 23/11/2011 - ^[GS]^
'Writes the "Imitate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Imitate)
End Sub

Public Sub WriteAlterGuildName(ByVal GuildName As String, ByVal newGuildName As String) ' 0.13.5
'***************************************************
'Author: Lex!
'Last Modification: 14/05/12
'Writes the "AlterGuildName" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterGuildName)
        
        Call .WriteASCIIString(GuildName)
        Call .WriteASCIIString(newGuildName)
    End With
End Sub

''
' Writes the "RecordAddObs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordAddObs(ByVal RecordIndex As Byte, ByVal Observation As String) ' 0.13.3
'***************************************************
'Author: Amraphen
'Last Modification: 23/11/2011 - ^[GS]^
'Writes the "RecordAddObs" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RecordAddObs)
        
        Call .WriteByte(RecordIndex)
        Call .WriteASCIIString(Observation)
    End With
End Sub

''
' Writes the "RecordAdd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordAdd(ByVal Nickname As String, ByVal Reason As String) ' 0.13.3
'***************************************************
'Author: Amraphen
'Last Modification: 23/11/2011 - ^[GS]^
'Writes the "RecordAdd" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RecordAdd)
        
        Call .WriteASCIIString(Nickname)
        Call .WriteASCIIString(Reason)
    End With
End Sub

''
' Writes the "RecordRemove" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordRemove(ByVal RecordIndex As Byte) ' 0.13.3
'***************************************************
'Author: Amraphen
'Last Modification: 23/11/2011 - ^[GS]^
'Writes the "RecordRemove" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RecordRemove)
        
        Call .WriteByte(RecordIndex)
    End With
End Sub

''
' Writes the "RecordListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordListRequest() ' 0.13.3
'***************************************************
'Author: Amraphen
'Last Modification: 23/11/2011 - ^[GS]^
'Writes the "RecordListRequest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RecordListRequest)
End Sub

''
' Writes the "RecordDetailsRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordDetailsRequest(ByVal RecordIndex As Byte) ' 0.13.3
'***************************************************
'Author: Amraphen
'Last Modification: 23/11/2011 - ^[GS]^
'Writes the "RecordDetailsRequest" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RecordDetailsRequest)
        
        Call .WriteByte(RecordIndex)
    End With
End Sub

''
' Handles the RecordList message.

Private Sub HandleRecordList() ' 0.13.3
'***************************************************
'Author: Amraphen
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim NumRecords As Byte
    Dim i As Long
    
    NumRecords = Buffer.ReadByte
    
    'Se limpia el ListBox y se agregan los usuarios
    frmPanelGm.lstUsers.Clear
    For i = 1 To NumRecords
        frmPanelGm.lstUsers.AddItem Buffer.ReadASCIIString
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the RecordDetails message.

Private Sub HandleRecordDetails() ' 0.13.3
'***************************************************
'Author: Amraphen
'Last Modification: 23/11/2011 - ^[GS]^
'
'***************************************************
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Dim tmpStr As String
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
       
    With frmPanelGm
        .txtCreador.Text = Buffer.ReadASCIIString
        .txtDescrip.Text = Buffer.ReadASCIIString
        
        'Status del pj
        If Buffer.ReadBoolean Then
            .lblEstado.ForeColor = vbGreen
            .lblEstado.Caption = "ONLINE"
        Else
            .lblEstado.ForeColor = vbRed
            .lblEstado.Caption = "OFFLINE"
        End If
        
        'IP del personaje
        tmpStr = Buffer.ReadASCIIString
        If LenB(tmpStr) Then
            .txtIP.Text = tmpStr
        Else
            .txtIP.Text = "Usuario offline"
        End If
        
        'Tiempo online
        tmpStr = Buffer.ReadASCIIString
        If LenB(tmpStr) Then
            .txtTimeOn.Text = tmpStr
        Else
            .txtTimeOn.Text = "Usuario offline"
        End If
        
        'Observaciones
        tmpStr = Buffer.ReadASCIIString
        If LenB(tmpStr) Then
            .txtObs.Text = tmpStr
        Else
            .txtObs.Text = "Sin observaciones"
        End If
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

Public Sub WriteDropObj(ByVal selInvObj As Byte, ByVal TargetX As Byte, ByVal TargetY As Byte, ByVal amount As Integer)
'***************************************************
'Author: maTih.-
'Last Modification: -
'Writes the "DropObj" message to the outgoing data buffer
'***************************************************

    With outgoingData
         .WriteByte ClientPacketID.DropObj
         .WriteByte selInvObj
         .WriteByte TargetX
         .WriteByte TargetY
         .WriteInteger amount
    End With

End Sub

''
' Writes the "Moveitem" message to the outgoing data buffer.
'
Public Sub WriteMoveItem(ByVal originalSlot As Integer, ByVal newSlot As Integer, ByVal moveType As eMoveType) ' 0.13.3
'***************************************************
'Author: Budi
'Last Modification: 23/11/2011 - ^[GS]^
'Writes the "MoveItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveItem)
        Call .WriteByte(originalSlot)
        Call .WriteByte(newSlot)
        Call .WriteByte(moveType)
    End With
End Sub

''
' Writes the "HigherAdminsMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to the other higher admins online.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHigherAdminsMessage(ByVal Message As String) ' 0.13.5
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 03/30/12
'Writes the "HigherAdminsMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.HigherAdminsMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "SearchObj" message to the outgoing data buffer.
'
' @param    NameObject
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSearchObj(ByVal NameObject As String) ' GSZAO
'***************************************************
'Author: ^[GS]^
'Last Modification: 02/08/2012 - ^[GS]^
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SearchObj)
        Call .WriteASCIIString(NameObject)
    End With
End Sub


''
' Writes the "SearchNpc" message to the outgoing data buffer.
'
' @param    NameNpc
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSearchNpc(ByVal NameNpc As String) ' GSZAO
'***************************************************
'Author: ^[GS]^
'Last Modification: 02/08/2012 - ^[GS]^
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SearchNpc)
        Call .WriteASCIIString(NameNpc)
    End With
End Sub

''
' Writes the "LluviaDeORO" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLluviaDeORO() ' GSZAO
'***************************************************
'Author: ^[GS]^
'Last Modification: 31/03/2013 - ^[GS]^
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.LluviaDeORO)
    End With
End Sub

''
' Handles the CreateParticle message.
Private Sub HandleCreateParticle()

        ' @ Crea FX/Particula sobre chars.

        Dim AttackerCharIndex As Integer
        Dim VictimCharIndex   As Integer
        Dim EffectIndex       As Integer
        Dim FXLoops           As Integer
        Dim ParticleCasteada  As Boolean

        With incomingData
        
                'Borra packetID
                .ReadByte
                
                'Carga data.
                AttackerCharIndex = .ReadInteger()
                VictimCharIndex = .ReadInteger()
                EffectIndex = .ReadInteger()
                FXLoops = .ReadInteger()
                
                If Not usaParticulas Then
                    Call SetCharacterFx(VictimCharIndex, EffectIndex, FXLoops)
                Else
                    ParticleCasteada = (Engine_UTOV_Particle(AttackerCharIndex, VictimCharIndex, EffectIndex) <> 0)

                    'Si quiere particulas, pero el hechizo no tiene particula
                    'Mostramos el fx.

                    If Not ParticleCasteada Then
                        Call SetCharacterFx(VictimCharIndex, EffectIndex, FXLoops)
                    End If
                End If
                
        End With

End Sub

Private Sub HandleInfoTorneo()


    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmTorneo
        OpcionTorneo = Buffer.ReadByte
    
        .lstParticipantes.Clear
        
        Participantes = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        Dim i As Long
        
        For i = 0 To UBound(Participantes) - 1
            Call .lstParticipantes.AddItem(Participantes(i))
        Next i
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
        
        .Show
    End With
    
End Sub

Public Sub WriteTorneoEvento(ByVal Opcion As Byte, Optional ByVal Participantes As Byte, Optional ByVal CaenItems As Byte)

    With outgoingData
        .WriteByte ClientPacketID.TorneoEvento
        
        .WriteByte Opcion
        
        If Opcion = 1 Then
            .WriteByte Participantes
            .WriteByte CaenItems
        End If
    End With
    
End Sub

Public Sub WritePedirInfoTorneo()
    With outgoingData
        .WriteByte ClientPacketID.TorneoEventoInfo
    End With
    
End Sub

Public Sub WriteQuest() ' GSZAO
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Escribe el paquete Quest al servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.Quest)
End Sub
 
Public Sub WriteQuestDetailsRequest(ByVal QuestSlot As Byte) ' GSZAO
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Escribe el paquete QuestDetailsRequest al servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.QuestDetailsRequest)
   
    Call outgoingData.WriteByte(QuestSlot)
End Sub
 
Public Sub WriteQuestAccept() ' GSZAO
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Escribe el paquete QuestAccept al servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.QuestAccept)
End Sub

Public Sub WriteQuestListRequest() ' GSZAO
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Escribe el paquete QuestListRequest al servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.QuestListRequest)
End Sub
 
Public Sub WriteQuestAbandon(ByVal QuestSlot As Byte) ' GSZAO
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Escribe el paquete QuestAbandon al servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el ID del paquete.
    Call outgoingData.WriteByte(ClientPacketID.QuestAbandon)
   
    'Escribe el Slot de Quest.
    Call outgoingData.WriteByte(QuestSlot)
End Sub

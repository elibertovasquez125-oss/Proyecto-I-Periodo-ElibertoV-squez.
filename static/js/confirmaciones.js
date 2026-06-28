// Muestra una ventana de confirmación en el navegador antes de borrar; devuelve true si el usuario acepta o false si cancela
function confirmar_eliminar() {
    return confirm(
        "¿Está seguro/a que desea eliminar al contacto? Esta acción no puede ser revertida."
    );
}

// Muestra una ventana de confirmación en el navegador antes de guardar cambios; devuelve true si el usuario acepta o false si cancela
function confirmar_actualizar() {
    return confirm(
        "¿Está seguro/a que desea actualizar la información del contacto? Esta acción no puede ser revertida."
    );
}
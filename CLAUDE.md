# Instrucciones para Claude

## Sincronización entre PCs (obligatorio)

Este proyecto se trabaja en 3 computadoras (PC1 = casa, PC2 = bodega, lap = laptop) y se sincroniza por GitHub. El usuario es nuevo en git: explícale en español y sin tecnicismos.

1. **Al empezar a trabajar en este proyecto** (antes de leer o modificar código): corre `git pull`. Si hay conflictos o cambios locales sin guardar, explícalo y pregunta antes de seguir.
2. **Después de cada cambio terminado** (una función, un arreglo, una compilación o un deploy que funcione): corre
   `git add -A`, luego `git commit -m "<descripción corta en español>"` y luego `git push`.
   No hagas commit de código a medias que no funcione.
3. Nunca subas archivos `.env` ni contraseñas (deben estar en `.gitignore`).
4. Si el push falla porque hay cambios nuevos en GitHub, haz `git pull` primero y vuelve a intentar.

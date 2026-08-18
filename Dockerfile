# Build stage: full dependency tree, compile src/ to dist/.
FROM node:24-slim AS build
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci
COPY tsconfig.json tsconfig.build.json ./
COPY src ./src
RUN npm run build

# Runtime stage. Debian -slim, never Alpine: @resvg/resvg-js-linux-x64-gnu declares
# libc: ["glibc"] and has no musl prebuild, so an Alpine base either fails to install
# the optional package or fails at require() time.
FROM node:24-slim AS runtime
ENV NODE_ENV=production
WORKDIR /app

# A second install from the lockfile rather than a COPY of the build stage's
# node_modules. This is what keeps dev dependencies out of the image, and it is what
# proves the platform-specific optional package still resolves without them.
COPY package.json package-lock.json ./
RUN npm ci --omit=dev && npm cache clean --force

COPY --from=build /app/dist ./dist

# Metadata only. Cloud Run ignores EXPOSE and sets PORT itself; the default lives in
# the SPECS table in src/config.ts and is not duplicated as ENV here.
EXPOSE 8080

USER node

# Exec form, so node is PID 1 and Cloud Run's SIGTERM reaches the handler in
# src/index.ts rather than a shell.
CMD ["node", "dist/index.js"]

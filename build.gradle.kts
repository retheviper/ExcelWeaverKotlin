plugins {
    kotlin("jvm") version "1.4.31"
    id("com.github.johnrengelman.shadow") version "6.1.0"
}

group = "com.retheviper"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
}

tasks.withType<org.jetbrains.kotlin.gradle.tasks.KotlinCompile>().configureEach {
    kotlinOptions {
        jvmTarget = "11"
    }
}

dependencies {
    implementation(kotlin("stdlib"))
    implementation(kotlin("stdlib-jdk8"))
    implementation("org.apache.poi:poi:5.0.0")
    implementation("org.apache.poi:poi-ooxml:5.0.0")
    implementation("io.kotlintest:kotlintest:2.0.7")
}

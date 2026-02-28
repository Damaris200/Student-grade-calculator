plugins {
    kotlin("jvm") version "2.0.0"
    application
}

group = "com.gradecalc"
version = "1.0.0"

repositories {
    mavenCentral()
}

dependencies {
    implementation("org.apache.poi:poi:5.2.5")
    implementation("org.apache.poi:poi-ooxml:5.2.5")
    implementation("org.apache.xmlbeans:xmlbeans:5.1.1")
    implementation("org.apache.commons:commons-compress:1.26.1")
    implementation("commons-io:commons-io:2.15.1")
    implementation("org.apache.logging.log4j:log4j-core:2.20.0")

}

application {
    mainClass.set("com.gradecalc.MainKt")
}

tasks.jar {
    manifest {
        attributes["Main-Class"] = "com.gradecalc.MainKt"
    }
    duplicatesStrategy = DuplicatesStrategy.EXCLUDE
    from(configurations.runtimeClasspath.get().map { if (it.isDirectory) it else zipTree(it) })
}
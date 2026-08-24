document.addEventListener("DOMContentLoaded", function () {

    /* =====================================================
       ELEMENTS
    ===================================================== */

    const loader = document.getElementById("loader");
    const intro = document.getElementById("introVideoSection");
    const video = document.getElementById("introVideo");
    const website = document.getElementById("website");
    const skipVideo = document.getElementById("skipVideo");

    const enrollModal =
        document.getElementById("enrollmentModal") ||
        document.getElementById("enrollModal");

    const facultyModal =
        document.getElementById("facultyModal");


    /* =====================================================
       LOADING → VIDEO → WEBSITE
    ===================================================== */

    function finishIntro() {

        if (video) {
            video.pause();
        }

        if (intro) {
            intro.classList.add("finished");
        }

        if (website) {
            website.classList.add("show");
        }

        document.body.classList.remove("intro-playing");
    }


    function playIntro() {

        if (!video) {
            finishIntro();
            return;
        }

        video.muted = true;
        video.volume = 0;

        const playPromise = video.play();

        if (playPromise !== undefined) {

            playPromise
                .then(function () {

                    console.log(
                        "Intro video started successfully."
                    );

                })
                .catch(function (error) {

                    console.log(
                        "Video autoplay blocked:",
                        error
                    );

                    /*
                     * Do not leave the user on a black screen.
                     * Show the website if autoplay is blocked.
                     */

                    setTimeout(function () {
                        finishIntro();
                    }, 1500);

                });
        }
    }


    /*
     * Start loading animation.
     */

    if (loader) {

        setTimeout(function () {

            loader.classList.add("hide");

            setTimeout(function () {

                if (intro) {

                    intro.classList.add("active");

                    document.body.classList.add(
                        "intro-playing"
                    );

                    playIntro();

                } else {

                    finishIntro();

                }

            }, 700);

        }, 2500);

    } else {

        /*
         * If loader does not exist,
         * directly start video.
         */

        if (intro) {

            intro.classList.add("active");

            playIntro();

        } else {

            finishIntro();

        }

    }


    /* =====================================================
       VIDEO ENDED
    ===================================================== */

    if (video) {

        video.addEventListener(
            "ended",
            function () {

                console.log(
                    "Intro video completed."
                );

                finishIntro();

            }
        );

        /*
         * If video cannot load.
         */

        video.addEventListener(
            "error",
            function () {

                console.log(
                    "Intro video could not be loaded."
                );

                setTimeout(function () {
                    finishIntro();
                }, 1000);

            }
        );

    }


    /* =====================================================
       SKIP VIDEO
    ===================================================== */

    if (skipVideo) {

        skipVideo.addEventListener(
            "click",
            function () {

                finishIntro();

            }
        );

    }


    /* =====================================================
       MOBILE MENU
    ===================================================== */

    const menuButton =
        document.querySelector(".menu-toggle");

    const navLinks =
        document.getElementById("navLinks");

    const mobileNav =
        document.querySelector(".mobile-nav");


    window.toggleMenu = function () {

        if (navLinks) {

            navLinks.classList.toggle("active");

        }

        if (mobileNav) {

            mobileNav.classList.toggle("open");

        }

    };


    if (menuButton) {

        menuButton.addEventListener(
            "click",
            function () {

                if (navLinks) {

                    navLinks.classList.toggle(
                        "active"
                    );

                }

                if (mobileNav) {

                    mobileNav.classList.toggle(
                        "open"
                    );

                }

            }
        );

    }


    /* =====================================================
       CLOSE MOBILE MENU AFTER CLICK
    ===================================================== */

    document
        .querySelectorAll(
            ".nav-links a, .mobile-nav a"
        )
        .forEach(function (link) {

            link.addEventListener(
                "click",
                function () {

                    if (navLinks) {
                        navLinks.classList.remove(
                            "active"
                        );
                    }

                    if (mobileNav) {
                        mobileNav.classList.remove(
                            "open"
                        );
                    }

                }
            );

        });


    /* =====================================================
       MODAL FUNCTIONS
    ===================================================== */

    function openModal(modal) {

        if (!modal) {
            return;
        }

        modal.classList.add("active");
        modal.classList.add("open");

        modal.setAttribute(
            "aria-hidden",
            "false"
        );

        document.body.classList.add(
            "modal-open"
        );

    }


    function closeModal(modal) {

        if (!modal) {
            return;
        }

        modal.classList.remove("active");
        modal.classList.remove("open");

        modal.setAttribute(
            "aria-hidden",
            "true"
        );

        if (
            !document.querySelector(
                ".modal.active, .modal.open"
            )
        ) {

            document.body.classList.remove(
                "modal-open"
            );

        }

    }


    /* =====================================================
       ENROLLMENT
    ===================================================== */

    window.openEnrollment = function (course = "") {

        openModal(enrollModal);

        if (course) {

            const courseSelect =
                document.getElementById(
                    "courseSelect"
                );

            if (courseSelect) {

                courseSelect.value = course;

            }

        }

    };


    window.closeEnrollment = function () {

        closeModal(enrollModal);

    };


    /*
     * Buttons using:
     * data-open-enroll
     */

    document
        .querySelectorAll(
            "[data-open-enroll]"
        )
        .forEach(function (button) {

            button.addEventListener(
                "click",
                function () {

                    const course =
                        button.dataset.course ||
                        "";

                    window.openEnrollment(
                        course
                    );

                }
            );

        });


    /*
     * Buttons using:
     * data-close-enroll
     */

    document
        .querySelectorAll(
            "[data-close-enroll]"
        )
        .forEach(function (button) {

            button.addEventListener(
                "click",
                function () {

                    closeModal(
                        enrollModal
                    );

                }
            );

        });


    /* =====================================================
       FACULTY
    ===================================================== */

    window.openFaculty = function (role = "") {

        openModal(facultyModal);

        if (role) {

            const roleSelect =
                document.getElementById(
                    "roleSelect"
                ) ||
                document.querySelector(
                    '#facultyModal select[name="role"]'
                );

            if (roleSelect) {

                roleSelect.value = role;

            }

        }

    };


    window.closeFaculty = function () {

        closeModal(facultyModal);

    };


    /*
     * Faculty open buttons
     */

    document
        .querySelectorAll(
            "[data-open-faculty]"
        )
        .forEach(function (button) {

            button.addEventListener(
                "click",
                function () {

                    const role =
                        button.dataset.role ||
                        "";

                    window.openFaculty(
                        role
                    );

                }
            );

        });


    /*
     * Faculty close buttons
     */

    document
        .querySelectorAll(
            "[data-close-faculty]"
        )
        .forEach(function (button) {

            button.addEventListener(
                "click",
                function () {

                    closeModal(
                        facultyModal
                    );

                }
            );

        });


    /* =====================================================
       FACULTY ROLE CARDS
    ===================================================== */

    document
        .querySelectorAll(".role")
        .forEach(function (role) {

            role.addEventListener(
                "click",
                function () {

                    const selectedRole =
                        role.dataset.role ||
                        "";

                    const roleSelect =
                        document.getElementById(
                            "roleSelect"
                        ) ||
                        document.querySelector(
                            '#facultyModal select[name="role"]'
                        );

                    if (roleSelect) {

                        roleSelect.value =
                            selectedRole;

                    }

                    openModal(
                        facultyModal
                    );

                }
            );

        });


    /* =====================================================
       CLOSE MODAL BY CLICKING OUTSIDE
    ===================================================== */

    window.addEventListener(
        "click",
        function (event) {

            if (
                enrollModal &&
                event.target === enrollModal
            ) {

                closeModal(
                    enrollModal
                );

            }

            if (
                facultyModal &&
                event.target === facultyModal
            ) {

                closeModal(
                    facultyModal
                );

            }

        }
    );


    /* =====================================================
       ESC KEY
    ===================================================== */

    document.addEventListener(
        "keydown",
        function (event) {

            if (event.key === "Escape") {

                closeModal(
                    enrollModal
                );

                closeModal(
                    facultyModal
                );

            }

        }
    );


    /* =====================================================
       COURSE SWITCHING
    ===================================================== */

    window.showCourse = function (course) {

        const entrance =
            document.getElementById(
                "entranceCourse"
            );

        const foundation =
            document.getElementById(
                "foundationCourse"
            );

        const tabs =
            document.querySelectorAll(
                ".pathway-tab"
            );


        if (!entrance || !foundation) {
            return;
        }


        tabs.forEach(function (tab) {

            tab.classList.remove(
                "active"
            );

        });


        if (course === "entrance") {

            entrance.style.display =
                "grid";

            foundation.style.display =
                "none";

            if (tabs[0]) {

                tabs[0].classList.add(
                    "active"
                );

            }

        } else {

            entrance.style.display =
                "none";

            foundation.style.display =
                "grid";

            if (tabs[1]) {

                tabs[1].classList.add(
                    "active"
                );

            }

        }

    };


    /* =====================================================
       SCROLL REVEAL ANIMATION
    ===================================================== */

    const revealElements =
        document.querySelectorAll(
            ".about-card, " +
            ".course-card, " +
            ".faculty-item, " +
            ".contact-card, " +
            ".address-card"
        );


    if (
        "IntersectionObserver" in window
    ) {

        const observer =
            new IntersectionObserver(
                function (entries) {

                    entries.forEach(
                        function (entry) {

                            if (
                                entry.isIntersecting
                            ) {

                                entry.target.style.opacity =
                                    "1";

                                entry.target.style.transform =
                                    "translateY(0)";

                                observer.unobserve(
                                    entry.target
                                );

                            }

                        }
                    );

                },
                {
                    threshold: 0.15
                }
            );


        revealElements.forEach(
            function (element) {

                element.style.opacity =
                    "0";

                element.style.transform =
                    "translateY(30px)";

                element.style.transition =
                    "opacity .8s ease, " +
                    "transform .8s ease";

                observer.observe(
                    element
                );

            }
        );

    } else {

        /*
         * Fallback for older browsers.
         */

        revealElements.forEach(
            function (element) {

                element.style.opacity =
                    "1";

                element.style.transform =
                    "translateY(0)";

            }
        );

    }


    /* =====================================================
       HERO REVEAL
    ===================================================== */

    document
        .querySelectorAll(".reveal")
        .forEach(function (element) {

            requestAnimationFrame(
                function () {

                    element.classList.add(
                        "visible"
                    );

                }
            );

        });


    /* =====================================================
       FLASH MESSAGES
    ===================================================== */

    const flashes =
        document.querySelectorAll(
            ".flash"
        );


    flashes.forEach(function (flash) {

        setTimeout(
            function () {

                flash.style.opacity =
                    "0";

                flash.style.transform =
                    "translateY(-10px)";

                setTimeout(
                    function () {

                        flash.remove();

                    },
                    500
                );

            },
            6000
        );

    });


    /* =====================================================
       STOP BODY SCROLL WHEN MODAL IS OPEN
    ===================================================== */

    const modalStyle =
        document.createElement("style");

    modalStyle.innerHTML = `

        body.modal-open {
            overflow: hidden;
        }

    `;

    document.head.appendChild(
        modalStyle
    );


    /* =====================================================
       NAVBAR SCROLL EFFECT
    ===================================================== */

    const header =
        document.querySelector(
            ".site-header"
        );


    window.addEventListener(
        "scroll",
        function () {

            if (!header) {
                return;
            }

            if (window.scrollY > 50) {

                header.classList.add(
                    "scrolled"
                );

            } else {

                header.classList.remove(
                    "scrolled"
                );

            }

        }
    );


    /* =====================================================
       CONSOLE MESSAGE
    ===================================================== */

    console.log(
        "VIJAY Educational Services website loaded successfully."
    );

});
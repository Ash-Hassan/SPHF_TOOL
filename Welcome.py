import streamlit as st
import base64

st.set_page_config(
    page_title="Hello",
    page_icon="👋",
    layout="wide"
)

def get_base64(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

def build_markup_for_logo(
    png_file,
    background_position="50% 10%",
    margin_top="10%",
    image_width="85%",
    image_height="",
):
    binary_string = get_base64(png_file)
    return """
            <style>
                [data-testid="stSidebarNav"] {
                    background-image: url("data:image/png;base64,%s");
                    background-repeat: no-repeat;
                    background-position: %s;
                    margin-top: %s;
                    background-size: %s %s;
                }
            </style>
            """ % (
        binary_string,
        background_position,
        margin_top,
        image_width,
        image_height,
    )


def add_logo(png_file):
    logo_markup = build_markup_for_logo(png_file)
    st.markdown(
        logo_markup,
        unsafe_allow_html=True,
    )

add_logo("Images/Logo.png")



# def set_background(png_file):
#     bin_str = get_base64(png_file)
#     page_bg_img = '''
#     <style>
#     .stApp {
#     background-image: url("data:image/png;base64,%s");
#     background-size: cover;
#     }
#     </style>
#     ''' % bin_str
#     st.markdown(page_bg_img, unsafe_allow_html=True)

# set_background('Images/Login.png')

st.image("Images/Consolidation Tool.png")


st.write("# Revolutionizing the Working World through Technology!")

st.sidebar.success("Select the function above.")

st.markdown(
    """
    #### Welcome to the EY automation tool ! 👋
    **V 1.0**

    **👈 Select a page from the sidebar** to get some work done.

    ### Get in touch:
    - Ernst & Young ~> [EY](https://www.ey.com/en_us/locations/pakistan)
    - Sindh People's Housing for Flood Affectees ~> [SPHF](http://www.sphf.gos.pk/)
    - Data Analyst & Developer ~> [The Data Team](https://www.linkedin.com/in/mhassan-zaheer/)

"""
)
